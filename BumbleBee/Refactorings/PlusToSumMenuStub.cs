using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;

namespace BumbleBee.Refactorings
{
    class PlusToSumMenuStub : RangeRefactoring
    {
        private static readonly OpToAggregate opToAggregate = new OpToAggregate();

        public override void Refactor(Range applyto)
        {
            opToAggregate.Refactor(applyto);
        }

        public override bool CanRefactor(Range applyto)
        {
            return opToAggregate.CanRefactor(applyto);
        }

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.SingleColumn;
    }
}
