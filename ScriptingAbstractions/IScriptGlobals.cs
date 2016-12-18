using System.Collections.Generic;

namespace ScriptingAbstractions
{
    public interface IScriptGlobals
    {
        IDictionary<string, object> Parameters {get;}
    }
}
