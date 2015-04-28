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

        /// <summary>
        /// What type of targets a refactoring can Apply to
        /// </summary>
        RangeType AppliesTo { get; }
    }



    public interface INodeRefactoring : IRangeRefactoring
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

    // Default implementation for INodeRefactoring methods
    public abstract class NodeRefactoring : INodeRefactoring
    {
        public virtual void Refactor(Range applyto)
        {
            var subject = (Range)applyto.Item[1, 1];
            subject.Formula = "=" + Refactor(Helper.Parse(applyto,Context.Empty).Node).Print();
        }

        public virtual bool CanRefactor(Range applyto)
        {
            return AppliesTo.Fits(applyto) && CanRefactor(Helper.Parse(applyto,Context.Empty).Node);
        }

        public RangeType AppliesTo { get { return RangeType.SingleCell; } }
        public abstract ParseTreeNode Refactor(ParseTreeNode applyto);
        public abstract bool CanRefactor(ParseTreeNode applyto);
    }
}
