using System;

namespace ScriptingAbstractions
{
    public interface IGlobalsFactory
    {
        object Create(AppDomain ExecutingDomain);
    }

    public interface IGlobalsFactory<out TGlobals> : IGlobalsFactory
        where TGlobals : class, IScriptGlobals
    {
        new TGlobals Create(AppDomain ExecutingDomain);
    }
}
