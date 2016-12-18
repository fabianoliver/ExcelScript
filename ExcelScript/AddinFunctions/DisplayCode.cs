using ExcelDna.Integration;
using ScriptingAbstractions;
using System;
using System.Linq;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(DisplayCode), Description = "Displays either the entire code, or specific lines from a script as given by the script handle. Can be useful for debugging.", IsVolatile = false)]
        public static object DisplayCode(
            [ExcelArgument(Name = "ScriptHandle", Description = "A stored handled to the script which shall be displayed")] string ScriptHandle,
            [ExcelArgument(Name = "DisplayLineNumbers", Description = "If true, prepends line numbers to each line of the code")] bool DisplayLineNumbers = false,
            [ExcelArgument(Name = "LineStart", Description = "Line number of the first line to be displayed (1-based). Optional.")] int LineStart = 0,
            [ExcelArgument(Name = "LineEnd", Description = "Line number of the last line to be displayed (1-based). Optional. If LineStart is given, but LineEnd is not, LineEnd will be = LineStart (i.e. will display exactly one line).")] int LineEnd = 0
            )
        {
            var script = GetFromStoreOrThrow<IScript>(ScriptHandle);

            var scriptLines = script
                .Code
                .Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

            if (!scriptLines.Any())
                return "";

            int LineNumDigits = (int)Math.Floor(Math.Log10(scriptLines.Count()) + 1);
            string LineNumFormat = "D" + LineNumDigits;

            int _LineStart = (LineStart < 1) ? 1 : LineStart;
            int _LineEnd = Math.Min(scriptLines.Length, (LineEnd < 1) ? ((LineStart < 1) ? scriptLines.Length : (int)LineStart) : (int)LineEnd);

            var result1d = scriptLines
                .Select((x, i) => new { LineNumber = i + 1, CodeLine = x })
                .Where(x => x.LineNumber >= _LineStart && x.LineNumber <= _LineEnd)
                .Select(x => DisplayLineNumbers ? x.LineNumber.ToString(LineNumFormat) + ":   " + x.CodeLine : x.CodeLine)
                .ToArray();

            string[,] result = new string[result1d.Length, 1];
            for (int i = 0; i < result1d.Length; i++)
            {
                result[i, 0] = result1d[i];
            }


            return result;
        }
    }
}
