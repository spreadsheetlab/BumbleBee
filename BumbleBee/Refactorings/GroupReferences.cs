using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn3.Refactorings
{
    class GroupReferences : IRangeRefactoring
    {
        public void Refactor(Range applyto)
        {
            throw new NotImplementedException();
        }

        public bool CanRefactor(Range applyto)
        {
            throw new NotImplementedException();
        }

        private const RangeType appliesTo = RangeType.Range;
        public RangeType AppliesTo { get { return appliesTo; } }

        /// <summary>
        /// List of functions on which multiple arguments act the same as a single ArrayAsArgument parameter.
        /// Basically these are all the functions that have only a single ArrayAsArgument parameter
        /// </summary>
        /// <example>
        /// SUM(A1,B5:B10,K9) is identical to SUM((A1,B5:B10,K9)) and thus belongs in the list
        /// SMALL(A1,B2) is not identical to  SMALL((A1,B2)) and thus doesn't belong in it
        /// </example>
        private ISet<String> functionsWhichCanBeGrouped = new HashSet<string>()
        {
            // Source: http://superuser.com/questions/447492/is-there-a-union-operator-in-excel
            "SUM(",
            "COUNT(",
            "COUNTA(",
            "COUNTBLANK(",
            "LARGE(",
            "MIN(",
            "MAX(",
            "AVERAGE(",
        };
    }
}
