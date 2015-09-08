using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;

namespace BumbleBee.Refactorings
{
    class ExtractFormulaMenuStub : RangeRefactoring
    {
        public override void Refactor(Range applyto)
        {
            Globals.BBAddIn.bbMenuRefactorings.extractFormulaTp.Child.init(applyto);
            Globals.BBAddIn.bbMenuRefactorings.extractFormulaCtp.Visible = true;
        }

        public override bool CanRefactor(Range applyto)
        {
            if (!base.CanRefactor(applyto)) return false;

            // Make sure all cells have the same R1C1
            var r1c1 = applyto.Cells.Cast<Range>().First().FormulaR1C1;
            return applyto.Cells.Cast<Range>().All(cell => cell.FormulaR1C1 == r1c1);
        }

        protected override RangeShape.Flags AppliesTo
        {
            get { return RangeShape.Flags.NonEmpty; }
        }
    }
}
