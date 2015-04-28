using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn3.Refactorings.Util
{
    public static class Helper
    {
        private static readonly ExcelFormulaParser parser = new ExcelFormulaParser();

        public static Range TopLeft(this Range r)
        {
            return (Range)r.Item[1, 1];
        }

        public static ParseTreeNode Parse(this string formula)
        {
            return parser.ParseToTree(formula).Root;
        }

        public static ContextNode Parse(this string formula, Context Ctx)
        {
            return new ContextNode(Ctx, Parse(formula));
        }

        private static bool isNumeric(string s)
        {
            double n;
            return double.TryParse(s, out n);
        }

        /// <param name="cell">A single cell</param>
        public static ContextNode Parse(Range cell, Context Ctx = null)
        {
            if (cell.Count != 1) throw new ArgumentException("Must be a single cell", "cell");
            string f = cell.Formula;
            string toParse =
                cell.HasFormula ? f.Substring(1)
                : isNumeric(f) ? f
                // Parse as text, replace single " with double "" to avoid breaking the escape sequence
                : String.Format("\"{0}\"", f.Replace("\"","\"\""));
            return Parse(toParse, Ctx ?? CreateContext(cell));
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

        // BUG: Workbooks aren't considered yet because the parser can't parse workbook names yet
        public static Name Find(this Names names, NamedRangeDef nr, bool ignoreWorkbook = true)
        {
            return names.Cast<Name>()
                .FirstOrDefault(name =>
                    nr.Name == name.Name
                 && nr.Worksheet == (name.Parent is Workbook ? "" : name.Parent.Name)
                 && (ignoreWorkbook || nr.Workbook == (name.Parent is Workbook ? name.Parent : name.Parent.Parent).Name)
                 );
        }

        public static string SheetAndAddress(this Range r)
        {
            return String.Format("{0}!{1}", r.Worksheet.Name, r.Address[false,false]);
        }
    }
}
