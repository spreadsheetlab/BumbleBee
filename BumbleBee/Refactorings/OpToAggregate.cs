using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Irony.Parsing;
using XLParser;

namespace BumbleBee.Refactorings
{
    /// <summary>
    /// Transforms a sequence of identical operators to the corresponding aggregate function.
    /// + to SUM, * to PRODUCT, & to CONCATENATE.
    /// </summary>
    public class OpToAggregate : FormulaRefactoring
    {
        // This class has no state, singleton instance
        public static readonly OpToAggregate Instance = new OpToAggregate();

        public override ParseTreeNode Refactor(ParseTreeNode applyto)
        {
            ParseTreeNode current;
            string opToRefactor;

            if (!CanRefactor(applyto, out current, out opToRefactor))
            {
                throw new ArgumentException("Cannot refactor this formula", nameof(applyto));
            }

            // Gather the arguments to the aggregate function in this list
            var arguments = new List<ParseTreeNode>();

            // Traverse the tree while we encounter the op
            while (current.MatchFunction(opToRefactor))
            {
                // All the ops we transform are left associative so right part of the tree can never be other ops.
                // Left part might be more ops which we can also put into the aggregate
                arguments.Add(current.ChildNodes[2]);
                current = current.ChildNodes[0].SkipToRelevant(false);
            }
            // Last left node does not contain any other ops, add it to the arguments
            arguments.Add(current);
            // Arguments were pushed in reverse order
            arguments.Reverse();

            
            // construct the new formula with the aggregate
            var newformula = $"{functions[opToRefactor]}({string.Join(",", arguments.Select(ExcelFormulaParser.Print))})";
            return newformula.Parse();
        }

        public override bool CanRefactor(ParseTreeNode applyto)
        {
            ParseTreeNode relevant;
            string opToRefactor;
            return CanRefactor(applyto, out relevant, out opToRefactor);
        }

        private static bool CanRefactor(ParseTreeNode applyto, out ParseTreeNode relevant, out string opToRefactor)
        {
            relevant = applyto.SkipToRelevant();
            opToRefactor = "";

            if (!relevant.IsBinaryNonReferenceOperation()) return false;
            opToRefactor = relevant.GetFunction();

            return functions.ContainsKey(opToRefactor);
        }

        private static readonly IReadOnlyDictionary<string, string> functions = new Dictionary<string, string>()
        {
            {"+", "SUM"},
            {"*", "PRODUCT"},
            {"&", "CONCATENATE"}
        };
    }
}
