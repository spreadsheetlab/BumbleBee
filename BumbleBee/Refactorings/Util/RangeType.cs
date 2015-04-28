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


    public static class RangeType
    {
        private const Flags Empty = 0;
        private const Flags Cell = Flags.SingleCell | Flags.Connected | Flags.SingleColumn | Flags.SingleRow;

        public static RangeType.Flags Type(this Range r)
        {
            // Check range is not null or empty
            if (r == null || r.Count == 0)
            {
                return Empty;
            }

            // Check if the range is a single cell
            if (r.Count == 1)
            {
                return Cell;
            }

            // We're sure we have multiple cells now
            Flags rt = Flags.MultipleCells;

            // Check if the range consists of multiple disconnected ranges
            if (r.Areas.Count == 1)
            {
                rt |= Flags.Connected;
            }


            int firstrow = (r.Item[1, 1] as Range).Row;
            if(r.Cells.Cast<Range>().All(x => x.Row == firstrow)) {
                rt |= Flags.SingleRow;
            }

            int firstcolumn = (r.Item[1, 1] as Range).Column;
            if(r.Cells.Cast<Range>().All(x => x.Column == firstcolumn)) {
                rt |= Flags.SingleColumn;
            }

            return rt;

            /** // Faster iterative version if performance becomes a problem
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

        public static bool Fits(this RangeType t, Range r)
        {
            return (RangeType(r) & t) != 0;
        }

        [Flags]
        public enum Flags
        {
            SingleCell = 1 << 1,
            MultipleCells = 1 << 2,
            Connected = 1 << 3,
            SingleRow = 1 << 4,
            SingleColumn = 1 << 5
        }
    }
}
