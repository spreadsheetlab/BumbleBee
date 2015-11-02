using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;
using XLParser;
using Infotron.Util;

namespace BumbleBee.Refactorings
{
    class AgregrateToConditionalAggregrate : RangeRefactoring
    {
        public override void Refactor(Range applyto)
        {
            if(!CanRefactor(applyto)) throw new ArgumentException("Cannot refactor this range");

            // Refactor + to SUM first so we can focus on that case
            if (_opToAggregate.CanRefactor(applyto))
            {
                _opToAggregate.Refactor(applyto);
            }

            var node = Helper.Parse(applyto);
            var fname = node.GetFunction();
            var fargs = node.GetFunctionArguments();

            PrefixInfo prefix;
            bool columnEqual;
            List<int> columns;
            bool rowEqual;
            List<int> rows;
            if (!checkRowAndColumns(node, out prefix, out columnEqual, out columns, out rowEqual, out rows))
            {
                throw new ArgumentException("All references cells must be in the same row or column");
            }

            if (columnEqual)
            {
                // TODO: Get the excel range representing the summed values. Maybe .precedents?
                // Shift it to first column, check if determiner column, go to next column. Continue until column is empty and a sufficient number has been tried.
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        private static bool IsSingleCellReference(ParseTreeNode arg)
        {
            arg = arg.SkipToRelevant(false);
            // Cell can have prefix
            return arg.ChildNodes[(arg.ChildNodes.Count == 1 ? 0 : 1)].Is(GrammarNames.Cell);
        }

        private readonly OpToAggregate _opToAggregate = new OpToAggregate();

        public override bool CanRefactor(Range applyto)
        {
            // Shape check
            if (!base.CanRefactor(applyto)) return false;

            var node = Helper.Parse(applyto);

            // Refactor op to aggregate first so we only have to handle that case
            if (_opToAggregate.CanRefactor(node))
            {
                node = _opToAggregate.Refactor(node);
            }

            node = node.SkipToRelevant(false);
            if (!node.IsNamedFunction()) return false;
            var funcname = node.GetFunction();
            if(!functions.Contains(funcname)) return false;

            PrefixInfo prefix;
            bool columnEqual;
            List<int> column;
            bool rowEqual;
            List<int> row;
            return checkRowAndColumns(node, out prefix, out columnEqual, out column, out rowEqual, out row);
        }

        private static bool checkRowAndColumns(ParseTreeNode node, out PrefixInfo prefix, out bool columnEqual, out List<int> columns, out bool rowEqual, out List<int> rows)
        {
            var fargs = node.GetFunctionArguments().Select(arg => arg.SkipToRelevant()).ToList();

            prefix = null;
            columnEqual = true;
            columns = new List<int>();
            rowEqual = true;
            rows = new List<int>();

            // Check if all arguments are single-cell references
            // And if all ar in the same column/row and prefix

            // Check first cell for initial values to compare
            if (!fargs[0].Is(GrammarNames.Reference)) return false;
            if (!(fargs[0].ChildNodes[fargs[0].ChildNodes.Count == 1 ? 0 : 1]).Is(GrammarNames.Cell)) return false;
            prefix = (fargs[0].ChildNodes.Count == 1 ? null : fargs[0].ChildNodes[0])?.GetPrefixInfo();
            var loc = new Location((fargs[0].ChildNodes[fargs[0].ChildNodes.Count == 1 ? 0 : 1]).Print());

            var column = loc.Column1;
            var row = loc.Row1;

            foreach (var refnode in fargs)
            {
                if (!refnode.Is(GrammarNames.Reference)) return false;
                var index = refnode.ChildNodes.Count == 1 ? 0 : 1;
                // Check if it is a cell
                if (!refnode.ChildNodes[index].Is(GrammarNames.Cell)) return false;
                // Check if all prefixes are equal
                if (index == 0 && prefix != null) return false;
                if (index == 1 && !refnode.ChildNodes[0].GetPrefixInfo().Equals(prefix)) return false;
                loc = new Location(refnode.ChildNodes[index].Print());
                // Add rows/columns to the list and check if they are equal
                if (columnEqual && column != loc.Column1) columnEqual = false;
                if (rowEqual && row != loc.Row1) rowEqual = false;
                if (!columnEqual && !rowEqual) return false;
                rows.Add(loc.Row1);
                columns.Add(loc.Column1);
            }
            return true;
        }

        private static readonly IReadOnlyList<string> functions = new [] { "AVERAGE", "COUNT","SUM" };

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.SingleCell;

        /// <summary>Tests whether this </summary>
        private static bool isArgumentArray(ParseTreeNode n)
        {
            if (!n.Is(GrammarNames.Arguments))
            {
                return false;
            }
            return false;
        }
    }
}
