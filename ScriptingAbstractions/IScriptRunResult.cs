using System;
using System.Collections.Generic;

namespace ScriptingAbstractions
{
    public interface IScriptRunResult
    {
        bool IsSuccess { get; }
        object ReturnValue { get; }
        IEnumerable<Exception> Errors { get; }
    }
}
