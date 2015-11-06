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
        private static readonly AgregrateToConditionalAggregrate aggrToCondAggr = new AgregrateToConditionalAggregrate();
        private static readonly GroupReferences groupReferences = new GroupReferences();

        public override void Refactor(Range applyto)
        {
            bool didsomething = false;
            if (opToAggregate.CanRefactor(applyto))
            {
                didsomething = true;
                opToAggregate.Refactor(applyto);
            }
            if (aggrToCondAggr.CanRefactor(applyto))
            {
                didsomething = true;
                aggrToCondAggr.Refactor(applyto);
            }
            if (didsomething)
            {
                groupReferences.Refactor(applyto);
            }
        }

        public override bool CanRefactor(Range applyto)
        {
            return opToAggregate.CanRefactor(applyto) || aggrToCondAggr.CanRefactor(applyto);
        }

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.NonEmpty;
    }
}
