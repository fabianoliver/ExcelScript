using ScriptingAbstractions;
using ScriptingAbstractions.Factory;

namespace RoslynScripting
{
    public sealed class ParseResult<TGlobals> : IParseResult<TGlobals>
       where TGlobals : class, IScriptGlobals
    {
        public IScript<TGlobals> Script { get; private set; }
        public string EntryMethodName { get; private set; }

        public ParseResult(IScript<TGlobals> Script, string EntryMethodName)
        {
            this.Script = Script;
            this.EntryMethodName = EntryMethodName;
        }
    }
}
