using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn3.Refactorings
{
    class AgregrateToConditionalAggregrate : RangeRefactoring
    {
        public override void Refactor(Range applyto)
        {
            throw new NotImplementedException();
        }

        public override bool CanRefactor(Range applyto)
        {
            // Shape check
            if (!base.CanRefactor(applyto)) return false;
            
            // Look for +'s/SUM's, COUNT and AVERAGE
            // TODO: Add +
            var ptn = Helper.Parse(applyto);
            // Get the count functions
            var candidates = ExcelFormulaParser.AllNodes(ptn)
                .Where(ExcelFormulaParser.IsFunction)
                // Only sums
                .Where(n =>
                {
                    switch (ExcelFormulaParser.GetFunction(n))
                    {
                        case "AVERAGE":
                        case "COUNT":
                        case "SUM":
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

        protected override RangeShape.Flags AppliesTo
        {
            // TODO: Extend to multiple cells on the same row & column
            get { return RangeShape.Flags.SingleCell; } //  | RangeShape.Flags.SingleColumn RangeShape.Flags.SingleRow
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
