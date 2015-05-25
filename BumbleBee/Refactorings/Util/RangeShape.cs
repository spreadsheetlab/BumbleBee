using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;

namespace ExcelAddIn3.Refactorings.Util
{
    /// <summary>
    /// Class that determines if the shape of a range has certain features, e.g. if it is nonempty or a single column.
    /// This can be used by refactorings, for example a refactoring might only be applicable to a single cell or connected range
    /// </summary>
    public static class RangeShape
    {
        private const Flags Empty = 0;
        private const Flags Cell = Flags.SingleCell | Flags.Connected | Flags.SingleColumn | Flags.SingleRow;

        public static Flags Shape(this Range r)
        {
            // Check range is not null or empty
            if (r == null || r.IsEmpty())
            {
                return Empty;
            }

            Flags rt = 0;

            // Check if there are nonempty cells
            rt |= HasContent(r);

            // Check if the range is a single cell
            if (r.Count == 1)
            {
                return rt | Cell;
            }

            // We're sure we have multiple cells now
            rt |= Flags.MultipleCells;

            // Check if the range consists of multiple disconnected ranges
            if (r.Areas.Count == 1)
            {
                rt |= Flags.Connected;
            }

            // Check if all cells are in a single row
            int firstrow = ((Range) r.Item[1, 1]).Row;
            if(r.Cells.Cast<Range>().All(x => x.Row == firstrow)) {
                rt |= Flags.SingleRow;
            }

            // Check if all cells are in a single column
            int firstcolumn = ((Range) r.Item[1, 1]).Column;
            if(r.Cells.Cast<Range>().All(x => x.Column == firstcolumn)) {
                rt |= Flags.SingleColumn;
            }

            return rt;

            /** // Faster iterative version (probably) if performance becomes a problem
            bool singlerow = true;
            bool singlecol = true;
            // Determine if cells are in a single column/row
            foreach (Range item in r.Cells)
            {
                singlerow = singlerow && item.Row == firstrow;
                singlecol = singlecol && item.Column == firstcol;

                if (!singlerow && !singlecol)
                {
                    break;
                }
            }

            if (singlerow) rt |= Flags.SingleRow;
            if (singlecol) rt |= Flags.SingleColumn;
            */
        }

        public static bool FitsShape(this Range r, Flags shape)
        {
            return (r.Shape() & shape) == shape;
        }

        public static bool Fits(this Flags shape, Range r)
        {
            return r.FitsShape(shape);
        }

        private static Flags HasContent(Range r)
        {
            return r.Cells.Cast<Range>().Any(cell => cell.Value2 != null && cell.Value2.ToString().Trim() != "") ? Flags.NonEmpty : 0;
        }

        [Flags]
        public enum Flags
        {
            SingleCell = 1 << 1,
            MultipleCells = 1 << 2,
            Connected = 1 << 3,
            SingleRow = 1 << 4,
            SingleColumn = 1 << 5,
            NonEmpty = 1 << 6,
        }
    }
}
