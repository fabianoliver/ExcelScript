using RoslynScripting.Internal.Marshalling;
using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;

namespace RoslynScripting.Internal
{
    internal interface IParseResult
    {
        string RefactoredCode { get; }
        string EntryMethodName { get; }
        IEnumerable<IParameter> Parameters { get; }
    }

    internal interface IScriptRunner
    {
        /// <summary>
        /// Does the script, when run, need to be recompiled?
        /// </summary>
        bool NeedsRecompilationFor(IParameter[] parameters, string scriptCode, ScriptingOptions Options);

        /// <summary>
        /// Asynchronously compiles and runs the script, returns its result
        /// </summary>
        RemoteTask<object> RunAsync(Func<AppDomain, IScriptGlobals> ScriptGlobalsFactory, IParameterValue[] parameters, string scriptCode, ScriptingOptions Options);

        RemoteTask<IParseResult> ParseAsync(string scriptCode, ScriptingOptions Options, Func<MethodInfo[], MethodInfo> EntryMethodSelector, Func<MethodInfo, IParameter[]> EntryMethodParameterFactory);

        RemoteTask<string> GetRtfFormattedCodeAsync(string scriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme);
        RemoteTask<FormattedText> GetFormattedCodeAsync(string scriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme);

        /*
        /// <summary>
        /// Asynchronously compiles and runs the script, returns its result
        /// </summary>
        /// <param name="submission">Instance of the compiler-generated submission object (e.g. Submission#0) that the script is/was run on</param>
        RemoteTask<IRunResult> RunAsync(Func<AppDomain, IScriptGlobals> ScriptGlobalsFactory, IParameterValue[] parameters, string scriptCode, ScriptingOptions Options, out object submission);*/
    }
}
