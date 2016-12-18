using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using System.Reflection;
using System.Threading.Tasks;

namespace RoslynScripting.Factory
{

    public class ScriptFactory : IScriptFactory
    {
        public IScript<TGlobals> Create<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory)
            where TGlobals : class, IScriptGlobals
        {
            return Create<TGlobals>(GlobalsFactory, null);
        }

        public IScript<TGlobals> Create<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions Options)
            where TGlobals : class, IScriptGlobals
        {
            return new Script<TGlobals>(GlobalsFactory, Options);
        }

        public void Inject<TGlobals>(IScript<TGlobals> Instance, Func<AppDomain, IScriptGlobals> GlobalsFactory)
            where TGlobals : class, IScriptGlobals
        {
            if (Instance == null)
                throw new ArgumentNullException(nameof(Instance));

            if (GlobalsFactory == null)
                throw new ArgumentNullException(nameof(GlobalsFactory));

            var instance = Instance as Script<TGlobals>;

            if(instance != null)
            {
                instance.GlobalsFactory = GlobalsFactory;
            }
        }

        public async Task<IParseResult<TGlobals>> ParseFromAsync<TGlobals>(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions ScriptingOptions, string code, Func<MethodInfo[], MethodInfo> EntryMethodSelector, Func<MethodInfo, IParameter[]> EntryMethodParameterFactory)
            where TGlobals : class, IScriptGlobals
        {
            if (GlobalsFactory == null)
                throw new ArgumentNullException(nameof(GlobalsFactory));

            if (code == null)
                throw new ArgumentNullException(nameof(code));

            if (EntryMethodSelector == null)
                throw new ArgumentNullException(nameof(EntryMethodSelector));

            if (EntryMethodParameterFactory == null)
                throw new ArgumentNullException(nameof(EntryMethodParameterFactory));

            if ((!EntryMethodSelector.Target?.GetType()?.IsSerializable) ?? false)
                throw new ArgumentException($"The {nameof(EntryMethodSelector)}'s target type must be serializable in order to use it accross AppDomains. This error often occurs if you reference any out-of-scope variables in your delegate, if its declaring type is not serializable.", nameof(EntryMethodSelector));

            if ((!EntryMethodParameterFactory.Target?.GetType()?.IsSerializable) ?? false)
                throw new ArgumentException($"The {nameof(EntryMethodParameterFactory)}'s target type must be serializable in order to use it accross AppDomains. This error often occurs if you reference any out-of-scope variables in your delegate, if its declaring type is not serializable.", nameof(EntryMethodParameterFactory));

            return await Script<TGlobals>.ParseFromAsync(GlobalsFactory, ScriptingOptions, code, EntryMethodSelector, EntryMethodParameterFactory);
        }



    }
}
