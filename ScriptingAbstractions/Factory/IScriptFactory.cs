using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ScriptingAbstractions.Factory
{
    public interface IParseResult<TGlobals>
        where TGlobals : class, IScriptGlobals
    {
        IScript<TGlobals> Script { get; }
        string EntryMethodName { get; }
    }

    public interface IScriptFactory
    {
        IScript<TGlobals> Create<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory)
            where TGlobals : class, IScriptGlobals;
       
        IScript<TGlobals> Create<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions Options)
            where TGlobals : class, IScriptGlobals;

        /// <summary>
        /// Injects dependencies into an existing instance. Can be used after deserialization.
        /// </summary>
        void Inject<TGlobals>(IScript<TGlobals> Instance, Func<AppDomain, IScriptGlobals> GlobalsFactory)
            where TGlobals : class, IScriptGlobals;

        Task<IParseResult<TGlobals>> ParseFromAsync<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions ScriptingOptions, string code, Func<MethodInfo[], MethodInfo> EntryMethodSelector, Func<MethodInfo, IParameter[]> EntryMethodParameterFactory)
            where TGlobals : class, IScriptGlobals;
    }
}
