using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using XLParser;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;
using stdole;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;

namespace BumbleBee.Refactorings.Util
{
    public static class Helper
    {
        public static Range TopLeft(this Range r)
        {
            return (Range)r.Item[1, 1];
        }

        public static bool IsEmpty(this Range r)
        {
            return r.Count == 0;
        }

        private static bool UseParseCache => true;

        // TODO: Replace with R1C1 cache and move references depending on memory usage and speed
        private static readonly ObjectCache formulaCache = UseParseCache ? new MemoryCache("FormulaCache") : null;

        public static ParseTreeNode Parse(this string formula)
        {
            if (UseParseCache && formulaCache.Contains(formula))
            {
                return (ParseTreeNode) formulaCache.Get(formula);
            }

            var parsed = ExcelFormulaParser.Parse(formula);
            if (UseParseCache)
            {
                formulaCache.Add(formula, parsed, new CacheItemPolicy());
            }
            return parsed;
        }

        public static ContextNode ParseCtx(this string formula, Context Ctx)
        {
            return new ContextNode(Ctx, Parse(formula));
        }

        private static bool isNumeric(string s)
        {
            double n;
            return double.TryParse(s, out n);
        }

        /// <param name="cell">A single cell</param>
        public static ContextNode ParseCtx(Range cell, Context Ctx = null)
        {
            return new ContextNode(Ctx ?? CreateContext(cell), Parse(cell));
        }

        public static ParseTreeNode Parse(Range cell)
        {
            if (cell.Count != 1) throw new ArgumentException("Must be a single cell", nameof(cell));
            string f = cell.Formula;
            string toParse =
                cell.HasFormula ? f.Substring(1)
                : isNumeric(f) ? f
                // Parse as text, replace single " with double "" to avoid breaking the escape sequence
                : $"\"{f.Replace("\"", "\"\"")}\"";
            return Parse(toParse);
        }

        public static Context CreateContext(this Range cell)
        {
            var definedIn = new ParserSheetReference(cell.Worksheet.Parent.Name,cell.Worksheet.Name);
            return new Context(definedIn, NamedRanges(cell.Application));
        }

        public static ISet<NamedRangeDef> NamedRanges(Application app)
        {
            var ret = new HashSet<NamedRangeDef>();

            foreach (Name name in app.Names)
            {
                var parentS = name.Parent as Worksheet;
                var sheet = parentS != null ? parentS.Name : "";
                var parentWb = (Workbook)(parentS != null ? parentS.Parent : name.Parent);
                ret.Add(new NamedRangeDef(parentWb.Name, sheet, name.Name));
            }

            return ret;
        }

        readonly static Regex cellAddressRegex = new Regex(@"\$?[A-Z]+\$?\d+", RegexOptions.IgnoreCase);
        public static bool isValidAddress(string targetAddress)
        {
            return cellAddressRegex.IsMatch(targetAddress);
        }

        public static Name Find(this Names names, NamedRangeDef nr)
        {
            return names.Cast<Name>()
                .FirstOrDefault(name =>
                {
                    Worksheet ws = null;
                    Workbook wb = null;
                    try
                    {
                        wb = name.Parent as Workbook;
                        if (wb != null)
                        {
                            return nr.Name == name.Name && nr.Workbook == wb.Name;
                        }
                        else
                        {
                            ws = name.Parent;
                            wb = ws.Parent;
                            return nr.Name == name.Name
                            && nr.Worksheet == ws.Name
                            && (nr.Workbook == wb.Name || nr.Workbook == "");
                        }
                    }
                    finally
                    {
                        ws.ReleaseCom();
                        wb.ReleaseCom();
                    }
                }
                );
        }

        public static string SheetAndAddress(this Range r)
        {
            var w = r.Worksheet;
            var name = w.Name;
            w.ReleaseCom();
            return $"{name}!{r.Address[false, false]}";
        }

        public static IEnumerable<Range> CellsToInspect(this Range r)
        {
            return r.Cells.Cast<Range>().Take(RangeRefactoring.MAX_CELLS);
        }

        /// <summary>
        /// Gives all unique formulas in this range (according to the R1C1 formula)
        /// </summary>
        /// <param name="cellsToExamine">Maximum number of cells to examine</param>
        public static IEnumerable<ParseTreeNode> UniqueFormulas(this Range r, int cellsToExamine = int.MaxValue)
        {
            var cells = r.Cells;
            var encountered = new HashSet<string>();
            var formulas = new List<ParseTreeNode>();
            var count = 0;
            foreach (ExcelRaw.Range cell in cells)
            {
                try
                {
                    if (count >= cellsToExamine) break;
                    string r1c1 = cell.FormulaR1C1;
                    if (!encountered.Contains(r1c1))
                    {
                        encountered.Add(r1c1);
                        formulas.Add(Parse(cell));
                    }
                    count++;
                }
                finally
                {
                    Marshal.ReleaseComObject(cell);
                }
            }
            Marshal.ReleaseComObject(cells);
            return formulas;
        }

        public static void ReleaseCom(this object o)
        {
            if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
        }
    }
}
