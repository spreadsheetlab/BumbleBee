using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using Infotron.Parsing;
using XLParser;

namespace BumbleBee.Refactorings
{
    public class InlineFormula : RangeRefactoring
    {

        /// <summary>
        /// Inline all cells in a range into their dependents
        /// </summary>
        /// <exception cref="AggregateException">If any cells could not be inlined, with as innerexceptions the individual errors.</exception>
        public override void Refactor(Range toInline)
        {
            var errors = new List<Exception>();
            foreach (var area in toInline.Areas.Cast<Range>())
            {
                var ctx = area.TopLeft().CreateContext();
                foreach (Range cell in area.Cells)
                {
                    try
                    {
                        if (cell.FitsShape(RangeShape.Flags.NonEmpty))
                        {
                            RefactorSingle(cell, ctx);
                        }
                    }
                    catch (Exception e)
                    {
                        errors.Add(e);
                    }
                }
            }
            if (errors.Count > 0)
            {
                throw new AggregateException(errors);
            }
        }

        protected override RangeShape.Flags AppliesTo
        {
            get { return RangeShape.Flags.NonEmpty; }
        }

        private static void RefactorSingle(Range toInline, Context toInlineCtx)
        {
            var dependencies = GetAllDirectDependents(toInline);
            if (dependencies.Count == 0)
            {
                throw new InvalidOperationException(String.Format("Cell {0} has no dependencies", toInline.SheetAndAddress()));
            }

            var toInlineAST = Helper.ParseCtx(toInline, toInlineCtx);
            //MessageBox.Show(toInlineFormula);
            var toInlineAddress = Helper.ParseCtx(toInline.Address[false, false], toInlineCtx);

            var errors = new List<Exception>();
            foreach (Range dependent in dependencies)
            {
                try
                {
                    //Debug.Print(dependent.Address[false,false,XlReferenceStyle.xlA1,true]);
                    var dependentAST = Helper.ParseCtx(dependent);
                    if (dependentAST.Node == null)
                    {
                        throw new InvalidOperationException(String.Format("Could not parse formula of {0}",
                            dependent.SheetAndAddress()));
                    }
                    // Check if the dependent has the cell in a range
                    //var ranges = RefactoringHelper.ContainsCellInRanges(toInlineAddress, dependentAST);
                    var ranges = toInlineAddress.CellContainedInRanges(dependentAST).ToList();
                    if (ranges.Any())
                    {
                        // TODO: Handle cell in range gracefully, e.g. by altering the range
                        throw new InvalidOperationException(
                            String.Format("{1} refers to cell in range '{0}'",
                                ranges.First().Print(), dependent.SheetAndAddress()));
                    }
                    // Check if the dependent has the cell in a named range

                    string range;
                    var nrs = dependentAST.NamedRanges
                        .Select(nr => toInline.Application.Names.Find(nr))
                        .Where(nr => nr != null);
                    if (IsInNamedRanges(toInline, nrs, out range))
                    {
                        throw new InvalidOperationException(
                            String.Format("{1} refers to cell in named range '{0}'.", range,
                                dependent.SheetAndAddress()));
                    }

                    // As a failsafe, check that the dependent cell at least refers to the to-inline cell.
                    // In case the above error conditions fail
                    if (!dependentAST.Contains(toInlineAddress))
                    {
                        throw new InvalidOperationException(String.Format("{1} refers to cell {0}, but it is unknown how.",dependent.SheetAndAddress(), toInline.Address[false,false]));
                    }

                    var newFormula = dependentAST.Replace(toInlineAddress, toInlineAST);
                    dependent.Formula = "=" + newFormula.Print();
                }
                catch (Exception e)
                {
                    errors.Add(e);
                }
            }

            if (errors.Count == 0)
            {
                toInline.Formula = "";
            }
            else
            {
                throw new AggregateException(
                    String.Format("Cell {0} could not be inlined in all dependents:\n{1}", toInline.Address[false, false], String.Join("\n",errors.Select(e => e.Message))),
                    errors);
            }
           
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
        private static ICollection<Range> GetAllDirectDependents(Range cell)
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
        private static bool IsInNamedRanges(Range cell, IEnumerable<Name> names, out string which)
        {
            var intersecting =
                names.FirstOrDefault(name =>
                {
                    var intersect = cell.Application.Intersect(cell, name.RefersToRange);
                    return intersect != null && intersect.Count > 0;
                });
            which = intersecting != null ? intersecting.Name : "";
            return intersecting != null;
        }

        /// <summary>
        /// Check if a cell is in the supplied named ranges
        /// </summary>
        private static bool IsInNamedRanges(Range cell, IEnumerable<Name> names)
        {
            string v;
            return IsInNamedRanges(cell, names, out v);
        }
    }
}