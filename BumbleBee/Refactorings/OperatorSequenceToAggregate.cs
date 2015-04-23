using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Infotron.Parsing;
using Irony.Parsing;
using P = Infotron.Parsing.ExcelFormulaParser;

namespace ExcelAddIn3.Refactorings
{
    /// <summary>
    /// Transforms a sequence of identical operators to the corresponding aggregate function.
    /// + to SUM, * to PRODUCT, & to CONCATENATE
    /// </summary>
    public class OperatorSequenceToAggregate : NodeRefactoring
    {
        public static readonly OperatorSequenceToAggregate Instance = new OperatorSequenceToAggregate();

        public override ParseTreeNode Refactor(ParseTreeNode applyto)
        {
            if(!CanRefactor(applyto)) throw new ArgumentException("Cannot refactor this formula", "applyto");
            var arguments = new List<ParseTreeNode>();

            var current = P.SkipToRevelantChildNodes(applyto);
            string opToRefactor = P.GetFunction(current);

            while (P.MatchFunction(current, opToRefactor))
            {
                // All the ops we transform are left associative so right part of the tree can never be other ops.
                // Left part might be more ops which we can also put into the aggregate
                arguments.Add(current.ChildNodes[2]);
                current = current.ChildNodes[0];
            }
            // Last left node does not contain any other ops, add it to the arguments
            arguments.Add(current);
            // construct the new formula with a SUM
            return Helper.Parse(String.Format("SUM({0})", String.Join(",", arguments.Select(P.Instance.Print))));
        }

        public override bool CanRefactor(ParseTreeNode applyto)
        {
            applyto = P.SkipToRevelantChildNodes(applyto);
            return P.IsFunction(applyto) && functions.ContainsKey(P.GetFunction(applyto));
        }

        private static readonly IReadOnlyDictionary<string, string> functions = new Dictionary<string, string>()
        {
            {GrammarNames.OpAddition, "SUM"},
            {GrammarNames.OpMultiplication, "PRODUCT"},
            {GrammarNames.OpConcatenation, "CONCATENATE"}
        };
    }
}
