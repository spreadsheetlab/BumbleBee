using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Infotron.Parsing;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

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
            // TODO: Replace by adjusting the AST instead of string replacement
            // This has problems with spaces and probably other cases
            string result = fullFormula.Replace(subFormula, targetAdress);
            if (!isValidFormula(result) && !isValidFormula(result.Substring(1)))
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
            ExcelFormulaParser P = new ExcelFormulaParser();
            try
            {
                return P.ParseToTree(formula) != null;
            }
            catch (InvalidDataException)
            {
                return false;
            }
        }

        public enum Direction
        {
            Left, Right, Up, Down, Fixed
        }

        /// <summary>
        /// Get all direct dependents of the given cell.
        /// </summary>
        /// <remarks>
        /// Contrary to Range.DirectDependents this also gives those in different sheets or workbooks.
        /// Won't give dependents in closed workbooks
        /// Won't give dependents of cells in protected worksheets
        /// Won't give dependents in hidden sheets
        /// Doesn't work with structured references because trace dependents doesn't work with them: https://social.msdn.microsoft.com/Forums/office/en-US/6fc03fe8-6805-45db-a556-35ebb3c4f396/in-vba-how-to-get-all-precedents-of-a-formula-containing-external-structure-references-ie?forum=exceldev
        /// Based on https://colinlegg.wordpress.com/2014/01/14/vba-determine-all-precedent-cells-a-nice-example-of-recursion/.
        /// </remarks>
        /// <returns>All dependent cells as a collection of ranges</returns>
        public static ICollection<Range> getAllDirectDependents(Range cell)
        {
            if (cell.Count > 1)
            {
                throw new ArgumentException("Range has more than one cell.");
            }

            String cellAddress = cell.Address[false, false, XlReferenceStyle.xlA1, true];

            // Disable updating the screen so the user doesn't see our trace arrows
            cell.Application.ScreenUpdating = false;

            var dependents = new List<Range>();

            cell.ShowDependents();
            // Unfortunately we don't know beforehand how many arrows and links there are, so we'll have to loop till we encounter a non-existing one
            bool checkNextArrowNumber = true;
            for(int arrow = 1; checkNextArrowNumber; arrow++)
            {
                checkNextArrowNumber = false;
                bool checkNextLink = true;
                for (int link = 1; checkNextLink; link++)
                {
                    checkNextLink = false;
                    try
                    {
                        Range dependent = cell.NavigateArrow(false, arrow, link);
                        // This still is a valid arrow, so check the next one
                        if (cellAddress != dependent.Address[false, false, XlReferenceStyle.xlA1, true])
                        {
                            checkNextArrowNumber = true;
                            checkNextLink = true;
                            dependents.Add(dependent);
                        }
                        // If you want to extend this to transitive dependencies, don't forget to do some circular reference detection
                    }
                    catch (COMException e)
                    {
                        if (e.ErrorCode == -2146827284 || e.Message == "NavigateArrow method of Range class failed") {
                            checkNextLink = false;
                        }
                        else
                        {
                            throw;
                        }
                        
                    }
                }
            }

            cell.ShowDependents(false);
            cell.Worksheet.ClearArrows();
            
            // Resume updating the screen
            cell.Application.ScreenUpdating = true;

            return dependents;
        }

        /// <summary>
        /// Check if a cell is in the supplied named ranges
        /// </summary>
        public static bool isInNamedRanges(Range cell, IEnumerable<String> ranges, out string which)
        {
            foreach (string range in ranges)
            {
                var excelRange = cell.Worksheet.Range[range];
                var intersect = cell.Application.Intersect(cell, excelRange);
                if (intersect != null && intersect.Count > 0)
                {
                    which = range;
                    return true;
                }
            }
            which = "";
            return false;
        }

        /// <summary>
        /// Check if a cell is in the supplied named ranges
        /// </summary>
        public static bool isInNamedRanges(Range cell, IEnumerable<String> ranges)
        {
            string v;
            return isInNamedRanges(cell, ranges, out v);
        }
        
    }
}
