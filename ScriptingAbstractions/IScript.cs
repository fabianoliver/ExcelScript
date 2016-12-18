using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ScriptingAbstractions
{
    public interface IScript : IHasDescription
    {
        /// <summary>
        /// Null in case of void
        /// </summary>
        Type ReturnType { get; set; }

        IList<IParameter> Parameters { get; }

        string Code { get; set; }

        IScriptRunResult Run(IEnumerable<IParameterValue> Parameters);
        Task<IScriptRunResult> RunAsync(IEnumerable<IParameterValue> Parameters);
    }

    public interface IScript<TGlobals> : IScript
        where TGlobals : class, IScriptGlobals
    {

    }
}
