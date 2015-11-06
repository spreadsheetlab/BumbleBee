using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;

namespace BumbleBee.Refactorings
{
    class NotImplementedRefactoring : RangeRefactoring
    {
        public override void Refactor(Range applyto)
        {
            throw new NotImplementedException();
        }

        public override bool CanRefactor(Range applyto)
        {
            return false;
        }

        protected override RangeShape.Flags AppliesTo
        {
            get { return 0; }
        }
    }
}
