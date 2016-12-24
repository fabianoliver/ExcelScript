using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using ExcelScript.Internal;
using ExcelScript.NinjectModules;
using ExcelScript.Registration;
using Ninject;
using Ninject.Extensions.Logging;
using Ninject.Extensions.Logging.Log4net;
using Ninject.Modules;
using ObjectStorage.Abstractions;
using RoslynScriptGlobals;
using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript
{
    public partial class ExcelScriptAddin : IExcelAddIn
    {
        private const string FunctionPrefix = "ExcelScript.";
        private const string FunctionReturn_ErrorPrefix = "#ERROR: ";

        private static IObjectStore m_ObjectStore;
        internal static ILogger m_Log;

        private static Func<AppDomain, IScriptGlobals> m_GlobalsFactory = GlobalsFactory;
        private static IScriptFactory m_ScriptFactory;
        private static IParameterFactory m_ParameterFactory;
        private static DirtyRangeFlagger m_DirtyRangeFlagger;
   
        public void AutoOpen()
        {
            try
            {
                SetupExceptions();
                SetupRegistration();
                SetupKernel();
                IntelliSenseServer.Register();

                m_Log.Info("ExcelScript Addin loaded");
            } catch (Exception ex)
            {
                throw;
            }
        }

        public void AutoClose()
        {
            if (m_DirtyRangeFlagger != null)
                m_DirtyRangeFlagger.Dispose();
        }

        private static IScriptGlobals GlobalsFactory(AppDomain TargetDomain)
        {
            if (TargetDomain != AppDomain.CurrentDomain)
                throw new InvalidOperationException($"Tried to create script globals for AppDomain {TargetDomain}, but was called in AppDomain {AppDomain.CurrentDomain}");

            var excelDnaUtilInitialize = typeof(ExcelDnaUtil).GetMethod("Initialize", BindingFlags.Static | BindingFlags.NonPublic);
            excelDnaUtilInitialize.Invoke(null, Array.Empty<object>());

            var typename = typeof(Globals).AssemblyQualifiedName;
            Type type = Type.GetType(typename);
            Func<Excel.Application> applicationFactory = () => new Excel.Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            var instance = (IScriptGlobals)Activator.CreateInstance(type, new object[] { applicationFactory });
            return instance;
        }

        private void SetupExceptions()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(HandleException);
        }

        private object HandleException(object ExceptionObject)
        {
            string message;

            var exception = ExceptionObject as Exception;
            if(exception != null)
            {
                message = ToErrorMessage(exception);
            } else if(ExceptionObject != null)
            {
                message = ExceptionObject.ToString();
            } else {
                message = "Unidentified Error";
            }

            return $"{FunctionReturn_ErrorPrefix}{message}";
        }

        private string ToErrorMessage(Exception ex)
        {
            Stack<Exception> handled = new Stack<Exception>();
            const int max_handled = 10;

            var result = new StringBuilder();

            while(ex != null && !handled.Contains(ex) && handled.Count < max_handled)
            {
                if (handled.Count > 0)
                    result.Append(" - ");

                result.Append(ex.Message);
                handled.Push(ex);
                ex = ex.InnerException;
            }

            return result.ToString();
        }

        private void SetupRegistration()
        {
            ProcessAndRegisterFunctions(ExcelRegistration.GetExcelFunctions());
        }

        private static void ProcessAndRegisterFunction(ExcelFunctionRegistration Registration)
        {
            ProcessAndRegisterFunctions(new ExcelFunctionRegistration[] { Registration });
        }

        private static void ProcessAndRegisterFunctions(IEnumerable<ExcelFunctionRegistration> Registrations)
        {
            var processedFunctions = ProcessFunctions(Registrations);
            processedFunctions.RegisterFunctions();
        }

        private static ExcelFunctionRegistration ProcessFunction(ExcelFunctionRegistration Registration)
        {
            return ProcessFunctions(new ExcelFunctionRegistration[] { Registration }).Single();
        }

        private static IEnumerable<ExcelFunctionRegistration> ProcessFunctions(IEnumerable<ExcelFunctionRegistration> Registrations)
        {
            // Set the Parameter Conversions before they are applied by the ProcessParameterConversions call below.
            // CONSIDER: We might change the registration to be an object...?
            var conversionConfig = GetParameterConversionConfig();
            var postAsyncReturnConfig = GetPostAsyncReturnConversionConfig();
            var functionHandlerConfig = GetFunctionExecutionHandlerConfig();


            return Registrations
                .UpdateRegistrationForRangeParameters()
                .ProcessParameterConversions(conversionConfig)
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                .ProcessParameterConversions(postAsyncReturnConfig)
                .ProcessParamsRegistrations()
                .ProcessFunctionExecutionHandlers(functionHandlerConfig);
        }

        static ParameterConversionConfiguration GetPostAsyncReturnConversionConfig()
        {
            return new ParameterConversionConfiguration()
                .AddReturnConversion((type, customAttributes) => type != typeof(object) ? null : ((Expression<Func<object, object>>)
                                                ((object returnValue) => returnValue.Equals(ExcelError.ExcelErrorNA) ? ExcelError.ExcelErrorGettingData : returnValue)));
        }

        static ParameterConversionConfiguration GetParameterConversionConfig()
        {
            var paramConversionConfig = new ParameterConversionConfiguration()

                // Register the Standard Parameter Conversions (with the optional switch on how to treat references to empty cells)
                .AddParameterConversion(ParameterConversions.GetNullableConversion(treatEmptyAsMissing: false))
                .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: false))

                // From string to Type
                .AddParameterConversion((string input) => ToType(input))

            // This parameter conversion adds support for string[] parameters (by accepting object[] instead).
            // It uses the TypeConversion utility class defined in ExcelDna.Registration to get an object->string
            // conversion that is consist with Excel (in this case, Excel is called to do the conversion).
                .AddParameterConversion((object[] inputs) => inputs.Select(TypeConversion.ConvertToString).ToArray())
                .AddParameterConversion(RangeParameterConversion.ParameterConversion);

            return paramConversionConfig;
        }


        internal static Excel.Range FromExcelReferenceToRange(ExcelReference excelReference)
        {
            Excel.Application app = new Excel.Application(null, ExcelDnaUtil.Application);
            object address = XlCall.Excel(XlCall.xlfReftext, excelReference, true);
            Excel.Range range = app.Range(address);
            return range;
        }


        static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
        {
            return new FunctionExecutionConfiguration()
                .AddFunctionExecutionHandler(FunctionLoggingHandler.LoggingHandlerSelector)
                .AddFunctionExecutionHandler(SuppressInDialogFunctionExecutionHandler.SuppressInDialogSelector);
        }

        private void SetupKernel()
        {
            using (var kernel = CreateKernel())
            {
                m_ObjectStore = kernel.Get<IObjectStore>();
                m_Log = kernel.Get<ILoggerFactory>().GetCurrentClassLogger();
                m_ScriptFactory = kernel.Get<IScriptFactory>();
                m_ParameterFactory = kernel.Get<IParameterFactory>();

                if (DirtyRangeFlagger.IsEnabled())
                    m_DirtyRangeFlagger = kernel.Get<DirtyRangeFlagger>();
            }

        }

        private IKernel CreateKernel()
        {
            ConfigureLog();
            var settings = new NinjectSettings { LoadExtensions = false };
            var kernel = new StandardKernel(settings, LoadModules());
           
            return kernel;
        }

        private INinjectModule[] LoadModules()
        {
            return new INinjectModule[]
            {
                new Log4NetModule(),
                new ScriptingModule()
            };
        }

        private void ConfigureLog()
        {
            if (!log4net.LogManager.GetRepository().Configured)
            {
                log4net.Config.XmlConfigurator.Configure();
            }
        }



        /// <summary>
        /// Gets the object from the store
        /// </summary>
        /// <typeparam name="T">Type of the object to retrieve</typeparam>
        /// <param name="HandleName">Name from which to retrieve the object</param>
        /// <returns>Object from the store</returns>
        /// <exception cref="ArgumentException">Thrown when the HandleName is (empty)</exception>
        /// <exception cref="Exception">Thrown when the given Handle is not found in the store</exception>
        private static T GetFromStoreOrThrow<T>(string HandleName, out IStoredObject<T> storedObj)
            where T : class
        {
            HandleName = HandleNames.GetNameFrom(HandleName);

            if (String.IsNullOrWhiteSpace(HandleName))
                throw new ArgumentException($"Invalid Handle name {HandleName}", nameof(HandleName));

            if (!m_ObjectStore.GetByName<T>(HandleName, out storedObj))
                throw new Exception($"No object {typeof(T).FullName} with handle named {HandleName} found");

            return storedObj.Object;
        }

        /// <summary>
        /// Gets the object from the store
        /// </summary>
        /// <typeparam name="T">Type of the object to retrieve</typeparam>
        /// <param name="HandleName">Name from which to retrieve the object</param>
        /// <returns>Object from the store</returns>
        /// <exception cref="ArgumentException">Thrown when the HandleName is (empty)</exception>
        /// <exception cref="Exception">Thrown when the given Handle is not found in the store</exception>
        private static T GetFromStoreOrThrow<T>(string HandleName)
            where T : class
        {
            IStoredObject<T> _unused;
            return GetFromStoreOrThrow(HandleName, out _unused);
        }

        /// <summary>
        /// Adds the given object to the object store, and returns its handle name (name:version).
        /// Throws exceptions if anything goes wrong.
        /// </summary>
        /// <typeparam name="T">Type of the object to add</typeparam>
        /// <param name="HandleName">Name under which the object shall be stored</param>
        /// <param name="obj">The object which shall be stored</param>
        /// <param name="storedObject">[out] the stored object</param>
        /// <returns>The Handle (name:version) under which the object was stored</returns>
        private static string AddOrUpdateInStoreOrThrow<T>(string HandleName, T obj, out IStoredObject<T> storedObject, Action OnUnregister = null)
        where T : class
        {
            // Remove any invalid parts from the input
            HandleName = HandleNames.GetNameFrom(HandleName);

            if (String.IsNullOrWhiteSpace(HandleName))
                throw new ArgumentException($"Invalid Handle name {HandleName}", nameof(HandleName));

            if (!m_ObjectStore.AddOrUpdate<T>(HandleName, obj, out storedObject, OnUnregister))
                throw new Exception("Could not add script to the object store");

            return HandleNames.ToHandle(storedObject);
        }



        /// <summary>
        /// Adds the given object to the object store, and returns its handle name (name:version).
        /// Throws exceptions if anything goes wrong.
        /// </summary>
        /// <typeparam name="T">Type of the object to add</typeparam>
        /// <param name="HandleName">Name under which the object shall be stored</param>
        /// <param name="obj">The object which shall be stored</param>
        /// <returns>The Handle (name:version) under which the object was stored</returns>
        private static string AddOrUpdateInStoreOrThrow<T>(string HandleName, T obj, Action OnUnregister = null)
        where T : class
        {
            IStoredObject<T> _unused;
            return AddOrUpdateInStoreOrThrow<T>(HandleName, obj, out _unused, OnUnregister);
        }

        private static HostingType ToHostingType(string input)
        {
            input = input.ToLower();

            switch (input)
            {
                case "global":
                    return HostingType.GlobalAppDomain;
                case "shared":
                    return HostingType.SharedSandboxAppDomain;
                case "individual":
                    return HostingType.IndividualScriptAppDomain;
                default:
                    throw new ArgumentException($"Expected input to be global, shared or individual. Was {input}.", "input");
            }
        }

        /// <summary>
        /// If <paramref name="obj"/> is not null and implements IDisposable, .Dispose() will be caled on it
        /// </summary>
        private static void TryDispose(object obj)
        {
            IDisposable disposable = obj as IDisposable;

            if (disposable != null)
                disposable.Dispose();
        }

        private static Excel.Range ToRange(ExcelReference excelReference)
        {
            Excel.Application app = new Excel.Application(null, ExcelDnaUtil.Application);
            Excel.Range refRange = app.Range(XlCall.Excel(XlCall.xlfReftext, excelReference, true));

            return refRange;
        }

        private static Type ToType(string TypeName)
        {
            TypeName = TypeName.ToLower();

            switch (TypeName)
            {
                case "int":
                case "integer":
                    return typeof(int);
                case "int[]":
                case "integer[]":
                    return typeof(int[]);
                case "int[,]":
                case "integer[,]":
                    return typeof(int[,]);

                case "str":
                case "string":
                    return typeof(string);
                case "str[]":
                case "string[]":
                    return typeof(string[]);
                case "str[,]":
                case "string[,]":
                    return typeof(string[,]);

                case "dbl":
                case "double":
                    return typeof(double);
                case "dbl[]":
                case "double[]":
                    return typeof(double[]);
                case "dbl[,]":
                case "double[,]":
                    return typeof(double[,]);

                case "flt":
                case "float":
                    return typeof(float);
                case "flt[]":
                case "float[]":
                    return typeof(float[]);
                case "flt[,]":
                case "float[,]":
                    return typeof(float[,]);

                case "obj":
                case "object":
                    return typeof(object);
                case "obj[]":
                case "object[]":
                    return typeof(object[]);
                case "obj[,]":
                case "object[,]":
                    return typeof(object[,]);

                case "type":
                    return typeof(Type);

                case "range":
                case "excel.range":
                    return typeof(Excel.Range);

                default:
                    throw new NotSupportedException($"Parameters type with name {TypeName} not recognized or supported. Valid types are int / string / double / float / uint / object.");
            }
        }
    }
}
