using ExcelDna.Integration;
using Irony.Parsing;
using Ninject.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using XLParser;
using Excel = NetOffice.ExcelApi;

namespace ExcelScript.Internal
{
    /// <summary>
    /// Experimental tool.
    /// When created, ensures that in all currently open workbooks, and all workboks that will be opened in the future,
    /// all cells calling a function of this Addin that is tagged with the <see cref="ManageDirtyFlagsAttribute"/> will be marked with <see cref="Excel.Range.Dirty"/> once.
    /// Enabling/disabling this can be done by setting app.config key TagDirtyMethodCalls = true.
    /// This feature could prove useful particularly in environments in manual calculation mode (with automatic recalculation, this may lead to all scripts in a workbook being run once opened, which may be undesirable)
    /// </summary>
    public class DirtyRangeFlagger : IDisposable
    {
        public static bool IsRecalculatingDirtyCells { get; private set; } = false;

        private ILogger Logger;
        private readonly string[] DirtyFunctionNames;
        private readonly Predicate<ParseTreeNode> DirtyFunctionParsePredicate;
        private Excel.Application Application;

        public static bool IsEnabled()
        {
            return Convert.ToBoolean(ConfigurationManager.AppSettings["TagDirtyMethodCalls"]);
        }

        public DirtyRangeFlagger(ILogger Logger)
        {
            Logger.Debug($"Initializing {typeof(DirtyRangeFlagger).Name}");

            this.Logger = Logger;
            this.Application = new Excel.Application(null, ExcelDnaUtil.Application);
            this.DirtyFunctionNames = GetDirtyFunctionNames();
            this.DirtyFunctionParsePredicate = GetDirtyFunctionParsePredicate();

            HookDirtyStateWatcher();

            Logger.Debug($"Initialized {typeof(DirtyRangeFlagger).Name}");
        }

        public void Dispose()
        {
            if (this.Application != null)
                this.Application.Dispose();
        }

        // We assume all addin functions are in the ExcelScriptAddin class,
        // and each function has its name explicitly set by ExcelFunctionAttribute.Name
        private string[] GetDirtyFunctionNames()
        {
            return typeof(ExcelScriptAddin)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(x => x.GetCustomAttribute<ManageDirtyFlagsAttribute>() != null)
            .Select(x => x.GetCustomAttribute<ExcelFunctionAttribute>())
            .Select(x => x.Name)
            .Select(x => x.ToUpper())
            .ToArray();
        }

        private Predicate<ParseTreeNode> GetDirtyFunctionParsePredicate()
        {
            return (tree) =>
                tree
                .SkipToRelevant()
                .AllNodes(GrammarNames.UDFunctionCall)
                .Select(ExcelFormulaParser.GetFunction)
                .Intersect(DirtyFunctionNames)
                .Any();
        }


        private void HookDirtyStateWatcher()
        {
            ExcelAsyncUtil.QueueAsMacro(
                delegate
                {
                    Excel.Application Application = new Excel.Application(null, ExcelDnaUtil.Application);
                    Application.WorkbookOpenEvent += App_WorkbookOpenEvent;

                    foreach (Excel.Workbook wb in Application.Workbooks)
                        SetDirtyFlagsIn(wb);
                });
        }

        private void App_WorkbookOpenEvent(Excel.Workbook Wb)
        {
            ExcelAsyncUtil.QueueAsMacro(
                delegate
                {
                    SetDirtyFlagsIn(Wb);
                });
        }


        private void SetDirtyFlagsIn(Excel.Workbook wb)
        {
            Logger.Debug($"Finding and setting dirty method cell candidates in workbook {wb?.Name}");
            IsRecalculatingDirtyCells = true;

            foreach (Excel.Worksheet sheet in wb.Sheets)
            {

                IEnumerable<Excel.Range> candidates = Enumerable.Empty<Excel.Range>();

                foreach (var function_name in DirtyFunctionNames)
                {
                    var new_candidates = FindAllCellsContainingFormulaCandidates(sheet, $"{function_name}(");

                    candidates = candidates.Concat(new_candidates);
                }

                candidates = candidates.Distinct();

                foreach (Excel.Range candidate in candidates)
                {
                    var formula = candidate.Formula as String ?? String.Empty;
                    bool formulaContanisCallToDirtyMethod = false;

                    try
                    {
                        var parse = ExcelFormulaParser.Parse(formula);
                        formulaContanisCallToDirtyMethod = DirtyFunctionParsePredicate.Invoke(parse);
                    }
                    catch (ArgumentException)
                    {
                        // parse failed. eg we're probably just parsing pure text or something.
                        formulaContanisCallToDirtyMethod = false;
                    }

                    if (formulaContanisCallToDirtyMethod)
                    {
                        Logger.Debug("Workbook {0}, Sheet {1}: Setting cell {2} to dirty calculation ({3})", wb?.Name, sheet?.Name, candidate?.Address, candidate?.Formula);
                        candidate.Dirty();
                    }
                    else
                    {
                        Logger.Debug("Workbook {0}, Sheet {1}: DISCARD cell {2} for dirty calculation ({3})", wb?.Name, sheet?.Name, candidate?.Address, candidate?.Formula);
                    }
                }
            }

            Logger.Debug("Done finding and setting dirty method cell candidates");
            IsRecalculatingDirtyCells = false;
        }

        private IEnumerable<Excel.Range> FindAllCellsContainingFormulaCandidates(Excel.Worksheet sheet, string pattern)
        {
            var result = new List<Excel.Range>();

            Excel.Range range = sheet.Cells.Find(
                    pattern,
                    after: null,
                    lookIn: Excel.Enums.XlFindLookIn.xlFormulas,
                    lookAt: Excel.Enums.XlLookAt.xlPart,
                    searchOrder: null,
                    searchDirection: null,
                    matchCase: false);

            if (range != null)
            {
                string firstAddress = range.Address;

                do
                {
                    result.Add(range);
                    range = sheet.Cells.FindNext(range);
                } while (range != null && range.Address != firstAddress);
            }

            return result;
        }
    }
}
