using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn3
{
    public static class RefactoringHelper
    {
        /// <summary>
        ///     Change part of a formula to a cell reference
        /// </summary>
        /// <exception cref="ArgumentException">If the subformula is not </exception>
        public static string replaceSubFormula(string fullFormula, string subFormula, string targetAdress)
        {
            if (!isValidFormula(subFormula))
            {
                throw new ArgumentException("Subformula is not a valid formula", "subFormula");
            }
            if (!isValidAddress(targetAdress))
            {
                throw new ArgumentException("Not a valid cell address", "targetAdress");
            }
            // Change cell to contain reference to new location
            // The string replace is a bit ugly, but I couldn't think of a case where it wouldn't work
            // as such using a transformationrule seems overpowered
            // TODO: Remove length check after formula is validated
            string result = (subFormula.Length > 0) ? fullFormula.Replace(subFormula, targetAdress) : fullFormula;
            if (!isValidFormula(result))
            {
                throw new InvalidOperationException(String.Format("After extraction new formula is not a valid formula: {0}", result));
            }
            return result;
        }

        readonly static Regex CellAddressRegex = new Regex(@"\$?[A-Z]+\$?\d+", RegexOptions.IgnoreCase);
        public static bool isValidAddress(string targetAddress)
        {
            return CellAddressRegex.IsMatch(targetAddress);
        }

        public static bool isValidFormula(string formula)
        {
            // TODO: Implement
            /*
                FSharpTransformationRule T = new FSharpTransformationRule();
                if (T.ParseToTree(value) == null)
                {
                    throw new ArgumentException("Not a valid formula.");
                }
                 */
            return true;
        }

        public enum Direction
        {
            Left, Right, Up, Down, Fixed
        }
    }
}
