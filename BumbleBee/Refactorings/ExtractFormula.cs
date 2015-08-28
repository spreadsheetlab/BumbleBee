using System;
using System.Linq;
using BumbleBee.Refactorings.Util;
using Infotron.Parsing;
using Infotron.Util;
using Microsoft.Office.Interop.Excel;

namespace BumbleBee.Refactorings
{
    public class ExtractFormula : RangeRefactoring
    {
        public Direction Dir { get; private set; }
        public ContextNode SubFormula { get; private set; }
        public Location To { get; private set; }

        public ExtractFormula(ContextNode subformula, Direction dir)
        {
            SubFormula = subformula;
            Dir = dir;
        }

        public ExtractFormula(ContextNode subformula, Location to)
        {
            SubFormula = subformula;
            To = to;
        }

        public class Direction : Tuple<int, int>
        {
            public enum DIR
            {
                Left,
                Right,
                Up,
                Down
            }

            public static readonly Direction Left = new Direction(DIR.Left, -1, 0);
            public static readonly Direction Right = new Direction(DIR.Right, 1, 0);
            public static readonly Direction Up = new Direction(DIR.Up, 0, 1);
            public static readonly Direction Down = new Direction(DIR.Down, 0, -1);

            public DIR Dir { get; private set; }
            public int x { get { return Item1; } }
            public int y { get { return Item2; } }

            public int ColOffset { get { return x; } }
            public int RowOffset { get { return -y; } }

            private Direction(DIR dir, int x, int y) : base(x, y)
            {
                Dir = dir;
            }

            public static bool operator ==(Direction a, Direction b) { return ReferenceEquals(a,null) ? ReferenceEquals(b, null) : a.Equals(b); }
            public static bool operator !=(Direction a, Direction b) { return !(a == b); }
        }

        public override void Refactor(Range applyto)
        {
            if (To != null)
            {
                Refactor(applyto, To, SubFormula);
            }
            else
            {
                Refactor(applyto, Dir, SubFormula);
            }
            
        }

        public override bool CanRefactor(Range applyto)
        {
            if (!base.CanRefactor(applyto)) return false;
            // Check if all cells contain the subformula
            return NotContainingSubformula(applyto, SubFormula) == null;
        }

        /// <summary>
        /// Extract the subformula's of a range in a certain direction
        /// </summary>
        public static void Refactor(Range applyto, Direction dir, ContextNode subformula)
        {
            var notContaining = NotContainingSubformula(applyto, subformula);
            if (notContaining != null)
            {
                throw new ArgumentException(String.Format((string) "Not all cells contain that subformula, for example: {0}", (object) notContaining.Address[false, false]));
            }

            if (dir == Direction.Left && applyto.TopLeft().Column == 1
                || dir == Direction.Up && applyto.TopLeft().Row == 1
                || applyto.Offset[dir.RowOffset, dir.ColOffset].Cells.Cast<Range>().Any(c => c.Value2 != null))
            {
                switch (dir.Dir)
                {
                    case Direction.DIR.Left:
                        applyto.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        //subformulaCells = applyto;
                        break;
                    case Direction.DIR.Right:
                        applyto.Offset[0, 1].Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        //subformulaCells = applyto;
                        break;
                    case Direction.DIR.Up:
                        applyto.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        //subformulaCells = applyto;
                        break;
                    case Direction.DIR.Down:
                        applyto.Offset[1,0].Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        //subformulaCells = applyto;
                        break;
                }
            }

            // Set all the cells which should contain the prototype
            Range subformulaCells = applyto.Offset[dir.RowOffset, dir.ColOffset];
            var prototype = subformulaCells.TopLeft();
            prototype.Formula = "=" + subformula.Print();
            var r1c1 = prototype.FormulaR1C1;
            foreach (Range subformulaCell in subformulaCells.Cells)
            {
                subformulaCell.FormulaR1C1 = r1c1;
            }

            // Set all the refactored cells
            foreach (var uniqueR1C1Group in applyto.Cells.Cast<Range>().GroupBy(c => c.FormulaR1C1))
            {
                prototype = uniqueR1C1Group.First();

                var parsed = Helper.ParseCtx(prototype);
                var targetAddr = parsed.Ctx.Parse(prototype.Offset[dir.RowOffset, dir.ColOffset].Address[false, false]);

                prototype.Formula = "=" + parsed.Replace(subformula, targetAddr).Print();
                r1c1 = prototype.FormulaR1C1;

                foreach (var cell in uniqueR1C1Group)
                {
                    cell.FormulaR1C1 = r1c1;
                }
            }
        }

        public static void Refactor(Range applyto, Location to, ContextNode subformula)
        {

            // Check if all cells contain the subformula
            var notContaining = NotContainingSubformula(applyto, subformula);
            if (notContaining != null)
            {
                throw new ArgumentException(String.Format((string)"Not all cells contain that subformula, for example: {0}", (object)notContaining.Address[false, false]));
            }

            Range target = applyto.Worksheet.Cells[to.Row1, to.Column1];
            if (target.Value2 != null && target.Value2 != "")
            {
                throw new ArgumentException(String.Format("Target cell {0} is not empty", to));
            }
            target.Formula = "=" + subformula.Print();
            var targetAddr = subformula.Ctx.Parse(to.Address());

            foreach (var uniqueR1C1Group in applyto.Cells.Cast<Range>().GroupBy(c => c.FormulaR1C1))
            {
                var prototype = uniqueR1C1Group.First();
                var parsed = Helper.ParseCtx(prototype);

                prototype.Formula = "=" + parsed.Replace(subformula, targetAddr).Print();
                var r1c1 = prototype.FormulaR1C1;

                foreach (var cell in uniqueR1C1Group)
                {
                    cell.FormulaR1C1 = r1c1;
                }
            }

            // TODO: Provide some sort of undo functionality if possible
            // Warning: You cannot allow Excel to undo the actions of a VSTO plugin, so that path is doomed to fail :(
            // Quote from page 176 from "Visual Studio Tools for Office 2007" by E. Carter:
            /* Undo in Excel
             *      Excel has an Undo method that can be used to undo the last few actions
             *      taken by the user. However, Excel does not support undoing actions taken
             *      by your code. As soon as your code touches the object model, Excel clears
             *      the undo history and it does not add any of the actions your code performs
             *      to the undo history.
             */
            // Best option would be a manual undo stack, but that still goes against user expectations:
            //      (will Ctrl+Z work?, cannot undo further than Add-in actions etc.)
            // To hook up on Excel's undo trigger there's [Application.OnUndo](https://msdn.microsoft.com/en-us/library/office/ff194135(v=office.15).aspx)
            //      but that still requires a VBA macro to be defined to undo the changes made by the addon.
        }

        private static Range NotContainingSubformula(Range applyto, ContextNode subformula)
        {
            var which = applyto.Cells.Cast<Range>()
                .GroupBy(c => (string) c.FormulaR1C1)
                .Select(group => new {parse = Helper.ParseCtx(@group.First(), subformula.Ctx), example = @group.First()})
                .FirstOrDefault(t => !t.parse.Contains(subformula));
            return which != null ? which.example : null;
        }

        protected override RangeShape.Flags AppliesTo
        {
            get { return RangeShape.Flags.NonEmpty; }
        }
    }
}