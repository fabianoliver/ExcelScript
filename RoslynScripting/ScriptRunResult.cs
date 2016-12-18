using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RoslynScripting
{
    public class ScriptRunResult : IScriptRunResult
    {
        public bool IsSuccess { get; private set; }
        public object ReturnValue { get; private set; }
        public IEnumerable<Exception> Errors { get; private set; }

        private ScriptRunResult(IEnumerable<Exception> exceptions)
        {
            this.IsSuccess = false;
            this.ReturnValue = null;
            this.Errors = exceptions;
        }

        private ScriptRunResult(Exception exception)
            : this(new Exception[] { exception })
        {
        }

        private ScriptRunResult(object ReturnValue)
        {
            this.IsSuccess = true;
            this.ReturnValue = ReturnValue;
            this.Errors = Enumerable.Empty<Exception>();
        }

        private ScriptRunResult()
        {

        }

        public static ScriptRunResult Success(object ReturnValue = null)
        {
            return new ScriptRunResult()
            {
                IsSuccess = true,
                ReturnValue = ReturnValue,
                Errors = Enumerable.Empty<Exception>()
            };
        }

        public static ScriptRunResult Failure(IEnumerable<Exception> Exceptions)
        {
            return new ScriptRunResult(Exceptions);
        }

        public static ScriptRunResult Failure(Exception exception)
        {
            return new ScriptRunResult(exception);
        }
    }
}
