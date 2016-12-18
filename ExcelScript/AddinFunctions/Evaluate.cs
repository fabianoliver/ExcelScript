using ExcelDna.Integration;
using ExcelScript.Registration;
using ScriptingAbstractions;
using System;
using System.Linq;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(Evaluate), Description = "Evaluates the given expression and returns its result", IsVolatile = false)]
        [SuppressInDialog]
        public static object Evaluate(
        [ExcelArgument(Description = "The code that is to be evaluated")] string[] code
        )
        {
            if (code == null)
                throw new ArgumentNullException(nameof(code));

            string _code = String.Join(Environment.NewLine, code);

            var script = CreateScript(XlScriptOptions.Default);

            script.Code = _code;
            script.ReturnType = typeof(object);
            var result = script.Run(Enumerable.Empty<IParameterValue>());

            if (result.IsSuccess)
            {
                return result.ReturnValue;
            }
            else
            {
                throw new AggregateException("There were errors running the script", result.Errors);
            }
        }
    }
}
