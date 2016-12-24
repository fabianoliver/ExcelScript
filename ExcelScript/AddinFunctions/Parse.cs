using ExcelDna.Integration;
using ExcelScript.Internal;
using ExcelScript.Registration;
using ObjectStorage.Abstractions;
using RoslynScriptGlobals;
using RoslynScripting;
using ScriptingAbstractions;
using System;
using System.Linq;
using System.Reflection;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(Parse), Description = "Creates a reusable script, and returns its Handle - which can then by invoked with the ExcelScript.Run method", IsVolatile = false)]
        [SuppressInDialog]
        [ManageDirtyFlags]
        public static object Parse(
          [ExcelArgument("The code for this script. Multiple lines are allowed.")] string[] code,
          [ExcelArgument("A handle to script option definitions as created by ExcelScript.Options.Create. Can be left blank to use defaults.")] string OptionsHandle = "",
          [ExcelArgument("Name for a handle under which the script object will be stored in memory, Can be ommitted, in which case the function name will be used.")] string HandleName = "")
        {
            if (code == null)
                throw new ArgumentNullException(nameof(code));

            string _code = String.Join(Environment.NewLine, code);

            // CAREFUL: This method may be run in a different AppDomain! Don't reference anything from outside of the delegate
            Func<MethodInfo[], MethodInfo> EntryMethodSelector = (methods) =>
            {
                MethodInfo result = null;

                // When only one method is available, we'll just return that one...
                if (methods.Count() == 1)
                {
                    result = methods.Single();
                }
                else
                {
                    // Otherwise, we'll try to find a (one) function marked with an [ExcelFunction] attribute
                    var markedMethods = methods
                        .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() != null);

                    if (markedMethods.Count() == 0)
                        throw new ArgumentException($"Your script defines {methods.Count()} methods; please mark one of them with the [ExcelFunction] attribute, to mark which function shall be used as an entry point to your script.");

                    if (markedMethods.Count() > 1)
                        throw new ArgumentException($"Please only mark one function with an [ExcelFunction] attribute; you have currently marked {markedMethods.Count()} functions: " + String.Join(", ", markedMethods.Select(x => x.Name)));

                    result = markedMethods.Single();
                }

                return result;
            };

            // CAREFUL: This method may be run in a different AppDomain! Don't reference anything from outside of the delegate
            Func<MethodInfo, IParameter[]> EntryMethodParameterFactory = (method =>
            {
                return method
                .GetParameters()
                .Select(x =>
                {
                    var excelArgumentAttribute = x.GetCustomAttribute<ExcelArgumentAttribute>();

                    return new Parameter
                    {
                        Name = (excelArgumentAttribute == null || String.IsNullOrWhiteSpace(excelArgumentAttribute.Name)) ? x.Name : excelArgumentAttribute.Name,
                        Type = x.ParameterType,
                        IsOptional = x.IsOptional,
                        DefaultValue = x.DefaultValue,
                        Description = (excelArgumentAttribute == null || String.IsNullOrWhiteSpace(excelArgumentAttribute.Description)) ? String.Empty : excelArgumentAttribute.Description
                    };
                })
                .ToArray();
            });

            // Create Script object
            XlScriptOptions xlScriptOptions = (String.IsNullOrEmpty(OptionsHandle)) ? XlScriptOptions.Default : GetFromStoreOrThrow<XlScriptOptions>(OptionsHandle);
            var options = CreateScriptingOptionsFrom(xlScriptOptions);
            options.References.Add(typeof(ExcelFunctionAttribute).Assembly);
            options.Imports.Add(typeof(ExcelFunctionAttribute).Namespace);

            var parseResult = m_ScriptFactory.ParseFromAsync<Globals>(m_GlobalsFactory, options, _code, EntryMethodSelector, EntryMethodParameterFactory).Result;
            var script = parseResult.Script;
            script.ReturnType = xlScriptOptions.ReturnType;

            if (String.IsNullOrWhiteSpace(HandleName))
                HandleName = parseResult.EntryMethodName;

            IStoredObject<IScript> storedScript;
            if (m_ObjectStore.GetByName<IScript>(HandleName, out storedScript) && storedScript.Object.GetHashCode() == script.GetHashCode())
            {
                return HandleNames.ToHandle(storedScript);
            }
            else
            {
                return AddOrUpdateInStoreOrThrow<IScript>(HandleName, script, () => TryDispose(script));
            }
        }
    }
}
