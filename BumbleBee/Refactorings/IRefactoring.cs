using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;


namespace ExcelAddIn3.Refactorings
{
    public interface IRangeRefactoring
    {
        /// <summary>
        /// Apply this refactoring to a specific range
        /// </summary>
        /// <exception cref="ArgumentException">If this refactoring cannot apply to the given range</exception>
        void Refactor(Range applyto);
        
        /// <summary>
        /// Test whether this refactoring can be applied to a range
        /// </summary>
        bool CanRefactor(Range applyto);
    }

    public abstract class RangeRefactoring : IRangeRefactoring
    {
        public abstract void Refactor(Range applyto);

        public virtual bool CanRefactor(Range applyto)
        {
            return applyto.FitsShape(AppliesTo);
        }

        /// <summary>
        /// What type of targets a refactoring can Apply to
        /// </summary>
        protected abstract RangeShape.Flags AppliesTo { get; }
    }


    public interface IFormulaRefactoring : IRangeRefactoring
    {
        /// <summary>
        /// Apply this refactoring to a specific ContextNode
        /// </summary>
        ParseTreeNode Refactor(ParseTreeNode applyto);

        /// <summary>
        /// Test whether this refactoring can be applied to this ContextNode
        /// </summary>
        bool CanRefactor(ParseTreeNode applyto);
    }

    // Default implementation for IFormulaRefactoring methods
    public abstract class FormulaRefactoring : RangeRefactoring, IFormulaRefactoring
    {
        public override void Refactor(Range applyto)
        {
            applyto.Formula = "=" + Refactor(Helper.Parse(applyto)).Print();
        }

        public override bool CanRefactor(Range applyto)
        {
            return applyto.FitsShape(AppliesTo) && CanRefactor(Helper.Parse(applyto));
        }

        protected override RangeShape.Flags AppliesTo { get { return RangeShape.Flags.SingleCell; } }

        public abstract ParseTreeNode Refactor(ParseTreeNode applyto);
        public abstract bool CanRefactor(ParseTreeNode applyto);
    }
}
