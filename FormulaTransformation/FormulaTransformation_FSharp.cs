using System;
using System.Collections.Generic;
using System.Linq;
using Irony.Parsing;
using Infotron.Util;
using Microsoft.FSharp.Collections;
using Infotron.Parsing;
using System.Diagnostics;
using FSharpEngine;


namespace Infotron.FSharpFormulaTransformation
{
    public class FSharpTransformationRule
    {
        public string Name;
        public ParseTreeNode from;
        public ParseTreeNode to;
        public double priority;

        public ParseTreeNode ParseToTree(string input) 
        {
            TransformationRuleGrammar C = new TransformationRuleGrammar();
            var parser = new Irony.Parsing.Parser(C);

            ParseTree P = parser.Parse(input);

            if (P.Status == ParseTreeStatus.Error)
            {
                Debug.WriteLine("Does not parse: " + input);
            }

            return P.Root;
        }


        public FSharpTransform.Formula CreateFSharpTree(ParseTreeNode input)
        {
            var termName = input.Term.Name;

            if (termName == GrammarNames.Reference ||
                termName == GrammarNames.Formula ||
                termName == GrammarNames.CellorRange ||
                termName == GrammarNames.Argument)
            {
                return CreateFSharpTree(input.ChildNodes.First());
            }
            else if (termName == GrammarNames.FunctionCall)
            {
                string FunctionName = "";
                List<FSharpTransform.Formula> arguments = new List<FSharpTransform.Formula>();

                foreach (var Argument in input.ChildNodes)
                {
                    if (ExcelFormulaParser.IsOperation(Argument) || (Argument.Term.Name == GrammarNames.Function))
                    {
                        FunctionName += Argument.ChildNodes.First().Token.ValueString;
                    }
                    else
                    {
                        foreach (var item in Argument.ChildNodes)
                        {
                            arguments.Add(CreateFSharpTree(item));
                        }
                        //compare shoiuld end in fix

                    }
                }

                FSharpList<FSharpTransform.Formula> Farguments = ConvertToFSharpList(arguments);
                return FSharpTransform.makeFormula(FunctionName, Farguments);
            }
            else if (termName == GrammarNames.Range)
            {
                if (input.ChildNodes.First().Term.ToString() == GrammarNames.DynamicRange)
                {
                    //get variables from dynamic cell
                    ParseTreeNode DynamicRange = input.ChildNodes.First();

                    ParseTreeNode VarExpression1 = DynamicRange.ChildNodes.ElementAt(0);
                    char Var1 = VarExpression1.Token.ValueString[0];

                    return FSharpTransform.makeDRange(Var1);
                }
                else
                {
                    ParseTreeNode Cell1 = input.ChildNodes.ElementAt(0);
                    ParseTreeNode Cell2 = input.ChildNodes.ElementAt(2);

                    FSharpTransform.SuperCell C1;
                    FSharpTransform.SuperCell C2;

                    if (Cell1.ChildNodes.First().Term.ToString() == GrammarNames.Cell)
                    {
                        string cell1Location = Cell1.ChildNodes.First().Token.ValueString;
                        Location L1 = new Location(cell1Location);
                        C1 = FSharpTransform.makeCell(L1.Column, L1.Row);
                    }
                    else
                    {
                        C1 = GetDynamicCell(Cell1);
                    }

                    if (Cell1.ChildNodes.First().Term.ToString() == GrammarNames.Cell)
                    {
                        string cell2Location = Cell2.ChildNodes.First().Token.ValueString;
                        Location L2 = new Location(cell2Location);
                        C2 = FSharpTransform.makeCell(L2.Column, L2.Row);
                    }
                    else
                    {
                        C2 = GetDynamicCell(Cell2);
                    }

                    return FSharpTransform.makeRange(C1, C2);
                }
            }
            else if (termName == GrammarNames.Cell)
            {
                if (input.ChildNodes.First().Term.ToString() == GrammarNames.DynamicCell)
                {
                    //get variables from dynamic cell
                    FSharpTransform.SuperCell x = GetDynamicCell(input);
                    return FSharpTransform.makeSuperCell(x);
                }
                else
                {
                    string cellLocation = input.ChildNodes.First().Token.ValueString;
                    Location L = new Location(cellLocation);
                    return FSharpTransform.makeSuperCell(FSharpTransform.makeCell(L.Column, L.Row));
                }
            }
            else if (termName == GrammarNames.Number)
            {
                string C = input.Token.ValueString;
                return FSharpTransform.makeConstant(C);
            }
            else if (termName == GrammarNames.Text)
            {
                string C3 = input.ChildNodes.First().Token.ValueString;
                return FSharpTransform.makeConstant("\"" + C3 + "\"");
            }
            else if (termName == GrammarNames.DynamicConstant)
            {
                //get variables from dynamic constanc
                char y = input.ChildNodes.First().Token.ValueString[0];
                return FSharpTransform.makeDArgument(y);
            }

            return null;
        }

        private static FSharpTransform.SuperCell GetDynamicCell(ParseTreeNode input)
        {
            ParseTreeNode DynamicCell = input.ChildNodes.First();

            ParseTreeNode VarExpression1 = DynamicCell.ChildNodes.ElementAt(0);
            ParseTreeNode VarExpression2 = DynamicCell.ChildNodes.ElementAt(1);

            char Var1;
            char Var2;
            char Var3;
            char Var4;

            if (VarExpression1.ChildNodes.Count == 1)
            {
                Var1 = VarExpression1.ChildNodes.First().Token.ValueString[0];
                Var2 = '0';
            }
            else
            {
                Var1 = VarExpression1.ChildNodes.First().ChildNodes.First().Token.ValueString[0];
                Var2 = VarExpression1.ChildNodes.ElementAt(2).ChildNodes.First().Token.ValueString[0];
            }

            if (VarExpression2.ChildNodes.Count == 1)
            {
                Var3 = VarExpression2.ChildNodes.First().Token.ValueString[0];
                Var4 = '0';
            }
            else
            {
                Var3 = VarExpression2.ChildNodes.First().ChildNodes.First().Token.ValueString[0];
                Var4 = VarExpression2.ChildNodes.ElementAt(2).ChildNodes.First().Token.ValueString[0];
            }

            FSharpTransform.SuperCell x = FSharpTransform.makeDCell(Var1, Var2, Var3, Var4);
            return x;
        }

        private FSharpList<FSharpTransform.Formula> ConvertToFSharpList(List<FSharpTransform.Formula> arguments)
        {
            if (arguments.Count == 1)
            {
                return new FSharpList<FSharpTransform.Formula>(arguments.First(), FSharpList<FSharpTransform.Formula>.Empty);
            }
            else
            {
                List<FSharpTransform.Formula> tail = new List<FSharpTransform.Formula>();

                for (int i = 1; i < arguments.Count; i++)
                {
                    tail.Add(arguments.ElementAt(i));
                }
                return new FSharpList<FSharpTransform.Formula>(arguments.First(), ConvertToFSharpList(tail));
            }
        }

        public bool CanBeAppliedonBool(string formula)
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree source = P.ParseToTree(formula);

            FSharpTransform.Formula FFrom = CreateFSharpTree(from);
            FSharpTransform.Formula FSource = CreateFSharpTree(source.Root);

            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public bool CanBeAppliedonBool(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = CreateFSharpTree(from);
            FSharpTransform.Formula FSource = CreateFSharpTree(source);
            
            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public FSharpMap<char, FSharpTransform.mapElement> CanBeAppliedonMap(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = CreateFSharpTree(from);
            FSharpTransform.Formula FSource = CreateFSharpTree(source);

            return FSharpTransform.CanBeAppliedonMap(FFrom, FSource);
        }


        public string ApplyOn(string formula)
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree source = P.ParseToTree(formula);

            FSharpTransform.Formula FFrom = CreateFSharpTree(from);
            FSharpTransform.Formula FTo = CreateFSharpTree(to);
            FSharpTransform.Formula FSource = CreateFSharpTree(source.Root);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return Print(result);
        }

        public FSharpTransform.Formula ApplyOn(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = CreateFSharpTree(from);
            FSharpTransform.Formula FTo = CreateFSharpTree(to);
            FSharpTransform.Formula FSource = CreateFSharpTree(source);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return result;
        }

        public string Print(FSharpTransform.Formula result)
        {
            if (result.IsS)
	        {
                var y = (FSharpTransform.Formula.S)result;
                var Cell = y.Item;

                var CCell = (FSharpTransform.SuperCell.C)Cell;
                string CellName = new Location(CCell.Item.Item1, CCell.Item.Item2).ToString();
                return CellName;               
	        }

            if (result.IsRange)
            {
                var y = (FSharpTransform.Formula.Range)result;
                var Cell1 = y.Item1;
                var Cell2 = y.Item2;
                
                string CellName1;
                string CellName2;

                if (Cell1.IsC)
                {
                    var CCell1 = (FSharpTransform.SuperCell.C)Cell1;
                    CellName1 = new Location(CCell1.Item.Item1, CCell1.Item.Item2).ToString();
                }
                else
                {
                    throw new ArgumentException("Unable to print dynamic tree");
                }

                if (Cell2.IsC)
                {
                    var CCell2 = (FSharpTransform.SuperCell.C)Cell2;
                    CellName2 = new Location(CCell2.Item.Item1, CCell2.Item.Item2).ToString();
                }
                else
                {
                    throw new ArgumentException("Unable to print dynamic tree");
                }
                
                return CellName1 + ":" + CellName2;
            }

            if (result.IsFunction)
            {
                var y = (FSharpTransform.Formula.Function)result;
                string FunctionName = y.Item1;

                if (FunctionName.Contains("(")) //it is a prefix function
                {
                    string Arguments = PrintArguments(y.Item2);
                    return FunctionName + Arguments + ")";
                }

                if (FunctionName == "")
	            {
                     string Arguments = PrintArguments(y.Item2);
                     return "(" + FunctionName + Arguments + ")";
	            }
                else //infix
                {
                    return Print(y.Item2.First()) + FunctionName + Print(y.Item2.ElementAt(1)); 
                }
           }

            if (result.IsRange)
            {
                var y = (FSharpTransform.Formula.Range)result;
                //string CellName1 = new Location(y.Item1.Item1, y.Item1.Item2).ToString();
                //string CellName2 = new Location(y.Item2.Item1, y.Item2.Item2).ToString();
                //return CellName1 + ":" + CellName2;
            }

            if (result.IsConstant)
            {
                var y = (FSharpTransform.Formula.Constant)result;
                return y.Item;
 
            }




            if (result.IsArgumentList)
            {
                var y = (FSharpTransform.Formula.ArgumentList)result;
                return PrintArguments(y.Item);
            }

            throw new ArgumentException("Unable to print dynamic tree");

        }

        private string PrintArguments(FSharpList<FSharpTransform.Formula> y)
        {
            string Arguments = "";
            foreach (FSharpTransform.Formula Argument in y)
            {
                Arguments += Print(Argument) + ",";
            }

            Arguments = RemoveFinalSymbol(Arguments);
            return Arguments;
        }

        private static string RemoveFirstSymbol(string input)
        {
            input = input.Substring(1, input.Length - 1);
            return input;
        }

        private static string RemoveFinalSymbol(string input)
        {
            input = input.Substring(0, input.Length - 1);
            return input;
        }
        
    }
}
