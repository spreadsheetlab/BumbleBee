using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using Infotron.Parsing;
using Infotron.Util;
using Irony.Parsing;

namespace ExcelAddIn3.Refactorings
{
    /// <summary>
    /// Group a set of references
    /// </summary>
    public class GroupReferences : FormulaRefactoring
    {
        private _Worksheet excel;
        public GroupReferences(_Worksheet excel)
        {
            this.excel = excel;
        }

        public GroupReferences(){}

        public override void Refactor(Range applyto)
        {
            excel = applyto.Worksheet;
            base.Refactor(applyto);
        }

        public override ParseTreeNode Refactor(ParseTreeNode applyto)
        {
            if (excel == null)
            {
                throw new InvalidOperationException("Must have reference to Excel worksheet to group references");
            }
            var targetFunctions = ExcelFormulaParser.AllNodes(applyto)
                .Where(IsTargetFunction);

            foreach (var function in targetFunctions)
            {
                var target = function;
                if(!varargsFunctions.Contains(ExcelFormulaParser.GetFunction(function)))
                {
                    // Not a varags functions, select only the arrayAsArgument argument
                    target = function.ChildNodes.First(x => x.Is(GrammarNames.ArrayAsArgument));
                }

                var arguments = target.ChildNodes
                    .First(x => x.Is(GrammarNames.Arguments))
                    .ChildNodes;

                var togroup = arguments
                    .Where(NodeCanBeGrouped);
                var toNotGroup = arguments
                    .Where(x => !NodeCanBeGrouped(x));

                var grouped = GroupTheReferences(togroup)
                    .OrderBy(x=>x) // Sort references alphabetically
                    .Select(x => x.Parse()); // Make them parsetreenodes again

                function.ChildNodes.Clear();
                function.ChildNodes.AddRange(toNotGroup);
                function.ChildNodes.AddRange(grouped);
            }

            return applyto;
        }

        public override bool CanRefactor(ParseTreeNode applyto)
        {
            return ExcelFormulaParser.AllNodes(applyto).Any(IsTargetFunction);
        }

        private static bool IsTargetFunction(ParseTreeNode node)
        {
            return
                    // Not interested in not-functions
                    ExcelFormulaParser.IsFullFunction(node)
                    // Or functions without arguments
                    && node.ChildNodes[1].ChildNodes.Any() 
                    && (varargsFunctions.Contains(ExcelFormulaParser.GetFunction(node))
                        // Functions have an arrayasargument parameter
                        || node.ChildNodes[1].ChildNodes.Any(n => n.Is(GrammarNames.ArrayAsArgument))
                       )
                   ;
        }

        private static bool NodeCanBeGrouped(ParseTreeNode node)
        {
            // can be grouped if the node is a reference
            var relevant = ExcelFormulaParser.SkipToRevelantChildNodes(node);
            return relevant.Is(GrammarNames.Reference)
                // no named ranges
                && !ExcelFormulaParser.AllNodes(node).Any(x=>x.Is(GrammarNames.NamedRange))
                // no vertical or horizontal ranges
                && !(relevant.ChildNodes[0].ChildNodes[0].Is(GrammarNames.Range) && relevant.ChildNodes[0].ChildNodes[0].ChildNodes.Count == 1);
        }

        /// <summary>
        /// Takes a list of references and return a grouped list of references
        /// </summary>
        private IEnumerable<string> GroupTheReferences(IEnumerable<ParseTreeNode> references)
        {
            var refs = references.Select(r => new {abs = checkAbsolute(r),reference = r.Print()}).ToList();
            var output = new List<string>();
            // We don't do anything with things that mix absolute and relative markers
            output.AddRange(refs.Where(x => x.abs.mixed).Select(x => x.reference));

            // Now make excel group everything, divided by absolute/relative row/columns
            var absoluteCategories = refs
                .Where(x => !x.abs.mixed)
                .GroupBy(x => new {colA = x.abs.colAbsolute, rowA = x.abs.rowAbsolute})
                ;
            foreach (var grouping in absoluteCategories)
            {
                var colA = grouping.Key.colA;
                var rowA = grouping.Key.rowA;
                var union = excel.Range[String.Join(",",grouping.Select(x => x.reference))];
                output.AddRange(union.Address[rowA,colA].Split(','));
            }

            return output;
        }

        /// <summary>
        /// Check if all cells references have the same row/col absolute type or if it's mixed
        /// </summary>
        private static Absolute checkAbsolute(ParseTreeNode reference)
        {
            var cells = ExcelFormulaParser.AllNodes(reference).Where(x => x.Is(GrammarNames.Cell));
            bool first = true;
            var a = new Absolute();
            var locs = cells.Select(cell => new Location(cell.Print()));
            foreach (var l in locs)
            {
                if (first)
                {
                    a.colAbsolute = l.ColumnFixed;
                    a.rowAbsolute = l.RowFixed;
                    a.mixed = false;
                    first = false;
                }
                else
                {
                    if (a.colAbsolute != l.ColumnFixed || a.rowAbsolute != l.RowFixed)
                    {
                        a.mixed = true;
                    }
                }
            }
            return a;
        }

        protected override RangeShape.Flags AppliesTo { get { return RangeShape.Flags.NonEmpty; } }

        /// <summary>
        /// List of functions on which multiple arguments act the same as a single ArrayAsArgument parameter.
        /// Basically these are all the functions that have only a single ArrayAsArgument parameter
        /// </summary>
        /// <example>
        /// SUM(A1,B5:B10,K9) is identical to SUM((A1,B5:B10,K9)) and thus belongs in the list
        /// SMALL(A1,B2) is not identical to  SMALL((A1,B2)) and thus doesn't belong in it
        /// </example>
        private static readonly ISet<String> varargsFunctions = new HashSet<string>()
        {
            // Source: http://superuser.com/questions/447492/is-there-a-union-operator-in-excel
            "SUM",
            "COUNT",
            "COUNTA",
            "COUNTBLANK",
            "LARGE",
            "MIN",
            "MAX",
            "AVERAGE",
        };

        private class Absolute
        {
            public bool colAbsolute = false;
            public bool rowAbsolute = false;
            // If it is mixed, e.g. the range $A1:A$7
            public bool mixed = false;
        }
    }
}
