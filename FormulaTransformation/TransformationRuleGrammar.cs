using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using XLParser;

namespace Infotron.FSharpFormulaTransformation
{
    // 0.2 - Based on XLParser
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

            var Disable = ToTerm("Disabled_rule_don't_type_this_and_if_you_do_you're_just_beginning_for_things_to_break", "disabled");

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
                | Number
                | VarExpression + InfixOp + VarExpression;

            // Constant arrays have the same syntax as dynamic cells, so disable them
            ConstantArray.Rule = Disable;
            MarkTransient(ConstantArray);
            DynamicCell.Rule = OpenCurlyParen + VarExpression + comma + VarExpression + CloseCurlyParen;


            // Structured references have the same syntax as dynamic constants, so disable them
            StructureReference.Rule = Disable;
            MarkTransient(StructureReference);
            DynamicConstant.Rule = EnclosedInBracketsToken;

            DynamicRange.Rule = OpenCurlyParen + LowLetter + CloseCurlyParen;

            // This solves reduce-reduce conflicts with multiple disabled rules
            var Disabled = new NonTerminal("DISABLED", Disable + ReduceHere());
            MarkTransient(Disabled);

            Reference.Rule = Reference.Rule | DynamicRange | DynamicCell | Disabled;

            Formula.Rule = Formula.Rule | DynamicConstant;

            #endregion
        }
    }
}
