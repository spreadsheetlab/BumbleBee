using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using Infotron.Parsing;
using Irony.Parsing;

namespace ExcelAddIn3.Refactorings
{
    /// <summary>
    /// Group a set of references
    /// </summary>
    class GroupReferences : RangeRefactoring
    {
        public override void Refactor(Range applyto)
        {
            var parsed = Helper.Parse(applyto);
            var targetFunctions = ExcelFormulaParser.AllNodes(parsed)
                .Where(IsTargetFunction);
        }

        public override bool CanRefactor(Range applyto)
        {
            return !applyto.IsEmpty();
        }

        private static bool IsTargetFunction(ParseTreeNode node)
        {
            return
                    // Not interested in not-functions
                    ExcelFormulaParser.IsFullFunction(node)
                    // Or functions without arguments
                    && node.ChildNodes[1].ChildNodes.Any() 
                    && varargsFunctions.Contains(ExcelFormulaParser.GetFunction(node))
                    // Functions have an arrayasargument parameter
                    || node.ChildNodes[1].ChildNodes.Any(n => n.Is(GrammarNames.ArrayAsArgument))
                   ;
        }

        /// <summary>
        /// Takes a list of references and return a grouped list of references
        /// </summary>
        private static IEnumerable<string> GroupTheReferences(IEnumerable<string> references, Application excel)
        {
            throw new NotImplementedException();
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
    }
}
