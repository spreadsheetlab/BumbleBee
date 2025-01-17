using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
using XLParser;

namespace BumbleBee.Refactorings
{
    public class InlineFormula : RangeRefactoring
    {
        #if DEBUG
            private static readonly bool DEBUG = true;
        #else
            private static readonly bool DEBUG = false;
        #endif

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
                    
                    catch (Exception e) when(!DEBUG)
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

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.NonEmpty;

        private static void RefactorSingle(Range toInline, Context toInlineCtx)
        {
            // If no AST to inline is provided, inline the toInline AST
            var toInlineAST = Helper.ParseCtx(toInline, toInlineCtx);

            // Gather dependencies
            var dependencies = GetAllDirectDependents(toInline);
            if (dependencies.Count == 0)
            {
                throw new InvalidOperationException($"Cell {toInline.SheetAndAddress()} has no dependencies");
            }

            
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
                        throw new InvalidOperationException($"Could not parse formula of {dependent.SheetAndAddress()}");
                    }
                    // Check if the dependent has the cell in a range
                    //var ranges = RefactoringHelper.ContainsCellInRanges(toInlineAddress, dependentAST);
                    var ranges = toInlineAddress.CellContainedInRanges(dependentAST);
                    if (ranges.Any())
                    {
                        var minimal = toInlineAST.Ctx.QualifyMinimal(ranges.First());
                        throw new InvalidOperationException($"{dependent.SheetAndAddress()} refers to cell in range '{minimal.Print()}'");
                    }
                    // Check if the dependent has the cell in a named range

                    string range;
                    var nrs = dependentAST.NamedRanges;
                    var excelnames = nrs
                        .Select(nr =>
                        {
                            var app = toInline.Application;
                            var names = app.Names;
                            var name = names.Find(nr);
                            names.ReleaseCom();
                            app.ReleaseCom();
                            return name;
                        })
                        .Where(nr => nr != null);
                    if (IsInNamedRanges(toInline, excelnames, out range))
                    {
                        throw new InvalidOperationException( $"{dependent.SheetAndAddress()} refers to cell in named range '{range}'.");
                    }

                    // As a failsafe, check that the dependent cell at least refers to the to-inline cell.
                    // In case the above error conditions fail
                    if (!dependentAST.Contains(toInlineAddress))
                    {
                        throw new InvalidOperationException($"{dependent.SheetAndAddress()} refers to cell {toInline.Address[false, false]}, but it is unknown how.");
                    }

                    var newFormula = dependentAST.Replace(toInlineAddress, toInlineAST);
                    try
                    {
                        dependent.Formula = "=" + newFormula.Print();
                    }
                    catch (COMException e)
                    {
                        throw new InvalidOperationException($"Refactoring produced invalid formula '={newFormula.Print()}' from original formula '{dependentAST.Print()}' for cell {dependent.SheetAndAddress()}", e);
                    }
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
        /// Get all direct dependents of the given cell, or first cell of the given range
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
        internal static ICollection<Range> GetAllDirectDependents(Range cell)
        {
            var firstCell = cell;//cell.Cells[1, 1];
            var cellAddress = firstCell.Address[false, false, XlReferenceStyle.xlA1, true];

            // Disable updating the screen so the user doesn't see our trace arrows
            cell.Application.ScreenUpdating = false;

            var dependents = new List<Range>();

            firstCell.ShowDependents();
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
                        Range dependent = firstCell.NavigateArrow(false, arrow, link);
                        // This still is a valid arrow, so check the next one
                        if (cellAddress != dependent.Address[false, false, XlReferenceStyle.xlA1, true])
                        {
                            checkNextArrowNumber = true;
                            checkNextLink = true;
                            dependents.Add(dependent);
                        }
                        //Marshal.ReleaseComObject(dependent);
                        // If you want to extend this to transitive dependencies, don't forget to do some circular reference detection
                    }
                    catch (COMException e) when (e.Message == "NavigateArrow method of Range class failed")
                    {
                        // Found the first invalid arrow
                        checkNextLink = false;
                    }
                }
            }

            firstCell.ShowDependents(false);
            cell.Worksheet.ClearArrows();
            
            // Resume updating the screen
            cell.Application.ScreenUpdating = true;

            //Marshal.ReleaseComObject(firstCell);
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