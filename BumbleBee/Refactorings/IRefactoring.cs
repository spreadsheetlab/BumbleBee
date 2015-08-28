using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Infotron.Parsing;
using Irony.Parsing;
using Microsoft.Office.Interop.Excel;


namespace BumbleBee.Refactorings
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

        /// <summary>
        /// Maximum number of cells to examine when checking if a refactoring applies.
        /// </summary>
        public const int MAX_CELLS = 64;
    }


    public interface IFormulaRefactoring : IRangeRefactoring
    {
        /// <summary>
        /// Apply this refactoring to a specific ContextNode. Allowed to change the original ParseTreeNode.
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
            //applyto.Formula = "=" + Refactor(Helper.Parse(applyto)).Print();
            foreach (Range cell in applyto.Cells)
            {
                var parsed = Helper.Parse(cell);
                if (CanRefactor(parsed))
                {
                    var refactored = Refactor(parsed);
                    try
                    {
                        cell.Formula = "=" + refactored.Print();
                    }
                    catch (COMException e)
                    {
                        throw new InvalidOperationException(String.Format("Refactoring produced invalid formula '{0}'", refactored.Print()), e);
                    }
                }
            }
        }

        public override bool CanRefactor(Range applyto)
        {
            return applyto.FitsShape(AppliesTo) && applyto.Cells.Cast<Range>().Any(cell => CanRefactor(Helper.Parse(cell)));
        }

        protected override RangeShape.Flags AppliesTo { get { return RangeShape.Flags.NonEmpty; } }

        public abstract ParseTreeNode Refactor(ParseTreeNode applyto);
        public abstract bool CanRefactor(ParseTreeNode applyto);
    }
}
