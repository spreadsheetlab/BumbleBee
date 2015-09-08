using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using XLParser;

namespace Infotron.FSharpFormulaTransformation
{
    [Language("Transformations", "0.2", "Grammar for Bumblebee transformation rules")]
    public class TransformationRuleGrammar : ExcelFormulaGrammar
    {
        public static class Names
        {
            public const string VarExpression = "Varexpression";
            public const string DynamicCell = "DynamicCell";
            public const string DynamicConstant = "DynamicConstant";
            public const string DynamicRange = "DynamicRange";
        }

        public TransformationRuleGrammar()
        {
            #region 1-Terminals - in PascalCase

            var LowLetter = new RegexBasedTerminal("LowLetter", "[a-z]");

            #endregion

            #region 2-NonTerminals

            var VarExpression = new NonTerminal(Names.VarExpression);
            var DynamicCell = new NonTerminal(Names.DynamicCell);
            var DynamicConstant = new NonTerminal(Names.DynamicConstant);
            var DynamicRange = new NonTerminal(Names.DynamicRange);

            #endregion

            #region 3-Rules

            VarExpression.Rule =
                LowLetter
                | comma
                | VarExpression + InfixOp + VarExpression;

            DynamicCell.Rule = OpenCurlyParen + VarExpression + comma + VarExpression + CloseCurlyParen;

            DynamicConstant.Rule = "[" + LowLetter + "]";

            DynamicRange.Rule = OpenCurlyParen + LowLetter + CloseCurlyParen;

            Cell.Rule =
                base.Cell.Rule
                | DynamicCell;

            Constant.Rule =
                base.Formula.Rule
                | DynamicConstant;

            #endregion
        }
    }
}
