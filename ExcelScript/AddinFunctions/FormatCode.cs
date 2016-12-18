using ExcelDna.Integration;
using RoslynScriptGlobals;
using RoslynScripting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript
{
    public partial class ExcelScriptAddin
    {
        [ExcelFunction(Name = FunctionPrefix + nameof(FormatCode), IsMacroType = true, Description = "Formats the code in the input range, and replaces the input ranges' data with the formatted code")]
        public static object FormatCode(
            [ExcelArgument(AllowReference = true, Description = "A cell reference to the range containing C# code")] object excelRef)
        {
            Excel.Range cellRef = ToRange((ExcelReference)excelRef);

            if (cellRef == null)
                throw new ArgumentException("Input was not an excel reference", nameof(excelRef));

            if (cellRef.Columns.Count != 1)
                throw new ArgumentException("The range must have exactly one column", nameof(excelRef));

            IEnumerable<string> codeLines = cellRef.Cells.Select(x => Convert.ToString(x.Value));

            // Remove all empty lines at the beginning and end
            codeLines = codeLines.SkipWhile(x => String.IsNullOrWhiteSpace(x))
                .Reverse()
                .SkipWhile(x => String.IsNullOrWhiteSpace(x))
                .Reverse();

            var code = String.Join(Environment.NewLine, codeLines);

            XlScriptOptions xlScriptOptions = XlScriptOptions.Default;
            var options = CreateScriptingOptionsFrom(xlScriptOptions);

            var formattedCode = Script<Globals>.GetFormattedCodeAsync(code, options).Result;

            if (formattedCode.TextLines.Count > cellRef.Rows.Count)
                throw new ArgumentException($"The formatted result has {formattedCode.TextLines.Count} lines, but the input range has only {cellRef.Rows.Count}; please expant your input range.", nameof(excelRef));

            for (int i = 0; i < cellRef.Rows.Count; i++)
            {
                Excel.Range target = cellRef[i + 1, 1];

                if (i >= formattedCode.TextLines.Count)
                {
                    // The target range is bigger than what we need; just clear out the unused cells
                    target.ClearContents();
                }
                else
                {
                    var line = formattedCode.TextLines.ElementAt(i);
                    target.Value = line.Text;

                    foreach (var linePart in line.Parts)
                    {
                        var characters = target.Characters(linePart.LineSpan.Start + 1, linePart.LineSpan.Length);
                        var textColor = linePart.TextFormat.TextColor;
                        var excelColor = Color.FromArgb(textColor.B, textColor.G, textColor.R);  // excel uses BGR format, so we convert RGB to BGR here
                        characters.Font.Color = excelColor.ToArgb();
                    }
                }
            }

            return $"({DateTime.Now.ToShortTimeString()}) Formatted";
        }
    }
}
