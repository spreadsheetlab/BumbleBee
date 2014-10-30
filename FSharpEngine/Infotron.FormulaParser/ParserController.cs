using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr.Runtime;
using Antlr.Runtime.Tree;

namespace Infotron.FormulaParser
{
    public class ParserController
    {
        public string Parse(string input)
        {
            Antlr.Runtime.ANTLRStringStream stream = new Antlr.Runtime.ANTLRStringStream(input);
            formulaLexer lexer = new formulaLexer(stream);
            CommonTokenStream tokens = new CommonTokenStream(lexer);
            formulaParser parser = new formulaParser(tokens);
            AstParserRuleReturnScope<CommonTree, CommonToken> result = parser.formula();
            String textual = result.Tree.ToStringTree();
            return textual;
        }
   
    }
}
