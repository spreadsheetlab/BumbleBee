using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Infotron.Parsing;
using Irony.Parsing;
using P = Infotron.Parsing.ExcelFormulaParser;

namespace ExcelAddIn3.Refactorings
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
                throw new ArgumentException("Cannot refactor this formula", "applyto");
            }

            // Gather the arguments to the aggregate function in this list
            var arguments = new List<ParseTreeNode>();

            // Traverse the tree while we encounter the op
            while (P.MatchFunction(current, opToRefactor))
            {
                // All the ops we transform are left associative so right part of the tree can never be other ops.
                // Left part might be more ops which we can also put into the aggregate
                arguments.Add(current.ChildNodes[2]);
                current = P.SkipToRevelantChildNodes(current.ChildNodes[0]);
            }
            // Last left node does not contain any other ops, add it to the arguments
            arguments.Add(current);
            // Arguments were pushed in reverse order
            arguments.Reverse();

            
            // construct the new formula with the aggregate
            var newformula = String.Format("{0}({1})", functions[opToRefactor], String.Join(",", arguments.Select(P.Instance.Print)));
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
            relevant = P.SkipToRevelantChildNodes(applyto);
            opToRefactor = "";

            if (!P.IsBinaryOperation(relevant)) return false;
            opToRefactor = P.GetFunction(relevant);

            return functions.ContainsKey(opToRefactor);
        }

        private static readonly IReadOnlyDictionary<string, string> functions = new Dictionary<string, string>()
        {
            {GrammarNames.OpAddition, "SUM"},
            {GrammarNames.OpMultiplication, "PRODUCT"},
            {GrammarNames.OpConcatenation, "CONCATENATE"}
        };
    }
}
