using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn3.Refactorings
{
    class AgregrateToConditionalAggregrate : IRangeRefactoring
    {
        public void Refactor(Range applyto)
        {
            throw new NotImplementedException();
        }

        public bool CanRefactor(Range applyto)
        {
            if (!AppliesTo.Fits(applyto))
            {
                return false;
            }
            // Look for +'s/SUM's, COUNT and AVERAGE
            // TODO: Add +
            ParseTreeNode ptn = Helper.Parse(applyto.Formula);
            // Get the count functions
            var candidates = ExcelFormulaParser.AllNodes(ptn)
                .Where(n => n.Is(GrammarNames.FunctionCall))
                // Only sums
                .Where(n =>
                {
                    switch (n.ChildNodes.First().ChildNodes.First().Token.ValueString)
                    {
                        case "AVERAGE(":
                        case "COUNT(":
                        case "SUM(":
                            return true;
                        default:
                            return false;
                    }
                });

            return candidates.Any(node =>
            {
                var args = node.ChildNodes[1];
                return args.ChildNodes.Count > 0 &&
                       (args.ChildNodes[0].Is(GrammarNames.ArrayAsArgument)
                       // If all are references 
                        || args.ChildNodes.All(arg => arg.ChildNodes[0].Is(GrammarNames.Reference))
                       );
            });
        }

        public RangeType AppliesTo
        {
            get { return RangeType.SingleColumn | RangeType.SingleRow; }
        }

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
