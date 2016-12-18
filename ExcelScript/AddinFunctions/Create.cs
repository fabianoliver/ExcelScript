using ExcelDna.Integration;
using ExcelScript.Internal;
using ExcelScript.Registration;
using ObjectStorage.Abstractions;
using RoslynScriptGlobals;
using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(Create), Description = "Creates a reusable script, and returns its Handle - which can then by invoked with the ExcelScript.Run method", IsVolatile = false)]
        [SuppressInDialog]
        [ManageDirtyFlags]
        public static object Create(
            [ExcelArgument("Name for a handle under which the script object will be stored in memory")] string HandleName,
            [ExcelArgument("The code for this script. Multiple lines are allowed.")] string[] code,
            [ExcelArgument("A handle to script option definitions as created by ExcelScript.Options.Create. Can be left blank to use defaults.")] string OptionsHandle = "",
            [ExcelArgument("An arbitrary number of handles to parameter definitions (as created by ExcelScript.Parameters.Create), which define input parameters for the script")] params string[] ParameterHandles)
        {
            if (code == null)
                throw new ArgumentNullException(nameof(code));

            string _code = String.Join(Environment.NewLine, code);

            // Get Options
            XlScriptOptions xlScriptOptions = (String.IsNullOrEmpty(OptionsHandle)) ? XlScriptOptions.Default : GetFromStoreOrThrow<XlScriptOptions>(OptionsHandle);

            // Create Script object
            var script = CreateScript(xlScriptOptions);
            script.Code = _code;
            script.ReturnType = xlScriptOptions.ReturnType;

            foreach (var parameterHandle in ParameterHandles)
            {
                var parameterHandleName = HandleNames.GetNameFrom(parameterHandle);
                IParameter storedParameter = GetFromStoreOrThrow<IParameter>(parameterHandleName);
                script.Parameters.Add(storedParameter);
            }

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

        private static IScript<Globals> CreateScript(XlScriptOptions xlScriptOptions)
        {
            var options = CreateScriptingOptionsFrom(xlScriptOptions);
            var script = m_ScriptFactory.Create<Globals>(m_GlobalsFactory, options);
            return script;
        }

        private static ScriptingOptions CreateScriptingOptionsFrom(XlScriptOptions xlScriptOptions)
        {
            var options = new ScriptingOptions();

            foreach (var reference in GetScriptDefaultReferences())
                options.References.Add(reference);

            foreach (var import in GetScriptDefaultImports())
                options.Imports.Add(import);

            options.HostingType = xlScriptOptions.HostingType;

            return options;
        }

        private static IEnumerable<Assembly> GetScriptDefaultReferences()
        {
            yield return typeof(string).Assembly;               // mscorlib
            yield return typeof(IComponent).Assembly;           // System
            yield return typeof(Enumerable).Assembly;           // System.Core
            yield return typeof(Globals).Assembly;              // script globals
            yield return typeof(Excel.Application).Assembly;    // NetOffice
        }


        private static IEnumerable<string> GetScriptDefaultImports()
        {
            yield return "System";
            yield return "System.Collections.Generic";
            yield return "System.Linq";
            yield return "NetOffice.ExcelApi";
            yield return "NetOffice";
        }

    }
}
