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
            string opToRefactor;
            var args = GetBinOpArguments(applyto, out opToRefactor);

            if (!functions.ContainsKey(opToRefactor))
            {
                throw new ArgumentException("Cannot refactor this formula", nameof(applyto));
            }

            // construct the new formula with the aggregate
            var newformula = $"{functions[opToRefactor]}({string.Join(",", args.Select(ExcelFormulaParser.Print))})";
            return newformula.Parse();
        }

        public override bool CanRefactor(ParseTreeNode applyto) => functions.ContainsKey(GetBinOp(applyto));

        private static readonly IReadOnlyDictionary<string, string> functions = new Dictionary<string, string>()
        {
            {"+", "SUM"},
            {"*", "PRODUCT"},
            {"&", "CONCATENATE"}
        };

        internal static string GetBinOp(ParseTreeNode node)
        {
            node = node.SkipToRelevant();

            return !node.IsBinaryNonReferenceOperation() ? "" : node.GetFunction();
        }

        /// <summary>
        /// Find all arguments of a binary operator. E.g. 1 + 2 + 3 returns [1,2,3]
        /// </summary>
        internal static List<ParseTreeNode> GetBinOpArguments(ParseTreeNode root, out string op)
        {
            var current = root.SkipToRelevant(false);
            op = GetBinOp(root);

            if(op == "") return new List<ParseTreeNode>();

            // Gather the arguments to the aggregate function in this list
            var arguments = new List<ParseTreeNode>();

            // Traverse the tree while we encounter the op
            while (current.MatchFunction(op))
            {
                // All the ops are left associative so right part of the tree can never be other ops.
                // Left part might be more ops which we can also put into the aggregate
                arguments.Add(current.ChildNodes[2]);
                current = current.ChildNodes[0].SkipToRelevant(false);
            }
            // Last left node does not contain any other ops, add it to the arguments
            arguments.Add(current);
            // Arguments were pushed in reverse order
            arguments.Reverse();

            return arguments;
        } 
    }
}
