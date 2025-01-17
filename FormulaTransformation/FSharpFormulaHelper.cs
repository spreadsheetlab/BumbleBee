﻿using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using FSharpEngine;
using Microsoft.FSharp.Collections;
using Infotron.Util;
using XLParser;

namespace Infotron.FSharpFormulaTransformation
{
    public static class FSharpFormulaHelper
    {
        private static readonly Parser parser = new Parser(new TransformationRuleGrammar());

        /// <summary>
        /// Parse a BumbleBee or Excel formula to a parse tree
        /// </summary>
        /// <param name="input">Formula without =</param>
        /// <returns>Parse tree on success or null on error.</returns>
        public static ParseTreeNode ParseToTree(string input)
        {
            var P = parser.Parse(input);
            return (P.Status == ParseTreeStatus.Error) ? null : P.Root;
        }

        /// <summary>
        /// Parse a BumbleBee or Excel formula to a F# AST
        /// </summary>
        /// <param name="input">Formula without =</param>
        /// <returns>AST on success or null on error.</returns>
        public static FSharpTransform.Formula createFSharpTree(string input)
        {
            var Ptree = ParseToTree(input);
            return (Ptree == null) ? null : CreateFSharpTree(Ptree);
        }

        /// <summary>
        /// Change a C# parse tree to a F# AST
        /// </summary>
        public static FSharpTransform.Formula CreateFSharpTree(this ParseTreeNode input)
        {
            if (input.IsParentheses())
            {
                return FSharpTransform.Formula.NewFunction("", ListModule.OfSeq(new [] { CreateFSharpTree(input.ChildNodes[0]) }));
            }

            input = input.SkipToRelevant();

            switch (input.Type())
            {
                case GrammarNames.FunctionCall:
                case GrammarNames.ReferenceFunctionCall:
                case GrammarNames.UDFunctionCall:
                    var fname = input.GetFunction() + (input.IsNamedFunction()?"(":"");
                    var args = ListModule.OfSeq(input.GetFunctionArguments().Select(CreateFSharpTree));
                    // Check for range
                    if (fname == ":")
                    {
                        return makeFSharpRange(input);
                    }
                    return FSharpTransform.makeFormula(fname, args);
                case GrammarNames.Reference:
                    // ignore prefix
                    return CreateFSharpTree(input.ChildNodes.Count == 1 ? input.ChildNodes[0] : input.ChildNodes[1]);
                case GrammarNames.Cell:
                    var L = new Location(input.Print());
                    return FSharpTransform.makeSuperCell(FSharpTransform.makeCell(L.Column, L.Row));
                case GrammarNames.NamedRange:
                    return FSharpTransform.makeNamedRange(input.Print());
                case TransformationRuleGrammar.Names.DynamicCell:
                    //get variables from dynamic cell
                    return FSharpTransform.makeSuperCell(GetDynamicCell(input));
                case TransformationRuleGrammar.Names.DynamicRange:
                    var letter = input // DynamicRange
                        .ChildNodes[0] // LowLetter
                        .Token.ValueString[0];
                    return FSharpTransform.makeDRange(letter);
                case GrammarNames.Constant:
                case GrammarNames.Number:
                case GrammarNames.Text:
                case GrammarNames.Bool:
                case GrammarNames.Error:
                case GrammarNames.RefError:
                    return FSharpTransform.makeConstant(input.Print());
                case TransformationRuleGrammar.Names.DynamicConstant:
                    return FSharpTransform.makeDArgument(input.ChildNodes[0].Token.ValueString[1]);
                default:
                    throw new ArgumentException($"Can't convert node type {input.Type()}", nameof(input));
            }
        }

        private static FSharpTransform.Formula makeFSharpRange(ParseTreeNode input)
        {
            ParseTreeNode Cell1 = input.ChildNodes[0];
            ParseTreeNode Cell2 = input.ChildNodes[2];

            FSharpTransform.SuperCell C1;
            FSharpTransform.SuperCell C2;

            if (Cell1.ChildNodes[0].Type() == GrammarNames.Cell)
            {
                string cell1Location = Cell1.ChildNodes[0].Print();
                Location L1 = new Location(cell1Location);
                C1 = FSharpTransform.makeCell(L1.Column, L1.Row);
            }
            else
            {
                C1 = GetDynamicCell(Cell1);
            }

            if (Cell1.ChildNodes[0].Type() == GrammarNames.Cell)
            {
                string cell2Location = Cell2.ChildNodes[0].Print();
                Location L2 = new Location(cell2Location);
                C2 = FSharpTransform.makeCell(L2.Column, L2.Row);
            }
            else
            {
                C2 = GetDynamicCell(Cell2);
            }

            return FSharpTransform.makeRange(C1, C2);
        }

        private static FSharpTransform.SuperCell GetDynamicCell(ParseTreeNode input)
        {
            ParseTreeNode DynamicCell = input.SkipToRelevant();

            ParseTreeNode VarExpression1 = DynamicCell.ChildNodes[0];
            ParseTreeNode VarExpression2 = DynamicCell.ChildNodes[2];

            char Var1;
            char Var2;
            char Var3;
            char Var4;

            if (VarExpression1.ChildNodes.Count == 1) {
                Var1 = Print(VarExpression1.ChildNodes[0])[0];
                Var2 = '0';
            } else {
                Var1 = Print(VarExpression1.ChildNodes[0])[0];
                Var2 = Print(VarExpression1.ChildNodes[2])[0];
            }

            if (VarExpression2.ChildNodes.Count == 1) {
                Var3 = Print(VarExpression2.ChildNodes[0])[0];
                Var4 = '0';
            } else {
                Var3 = Print(VarExpression2.ChildNodes[0])[0];
                Var4 = Print(VarExpression2.ChildNodes[2])[0];
            }

            FSharpTransform.SuperCell x = FSharpTransform.makeDCell(Var1, Var2, Var3, Var4);
            return x;
        }

        /// <summary>
        /// Print transformation rule grammar
        /// </summary>
        public static string Print(this ParseTreeNode input)
        {
            // For terminals, just print the token text
            if (input.Term is Terminal)
            {
                return input.Token.Text;
            }

            // (Lazy) enumerable for printed childs
            var childs = input.ChildNodes.Select(Print);
            // Concrete list when needed
            List<string> childsL;

            string ret;
            // Switch on nonterminals
            switch (input.Term.Name)
            {
                case TransformationRuleGrammar.Names.VarExpression:
                case TransformationRuleGrammar.Names.DynamicCell:
                case TransformationRuleGrammar.Names.DynamicConstant:
                case TransformationRuleGrammar.Names.DynamicRange:
                    return string.Join("", input.ChildNodes);

                case GrammarNames.Formula:
                    // Check if these are brackets, otherwise print first child
                    return input.IsParentheses() ? $"({childs.First()})" : childs.First();

                case GrammarNames.FunctionCall:
                case GrammarNames.ReferenceFunctionCall:
                case GrammarNames.UDFunctionCall:
                    childsL = childs.ToList();

                    if (input.IsNamedFunction())
                    {
                        return string.Join("", childsL) + ")";
                    }

                    if (input.IsBinaryOperation())
                    {
                        // format string for "normal" binary operation
                        string format = "{0} {1} {2}";
                        if (input.IsIntersection())
                        {
                            format = "{0} {2}";
                        }
                        else if (input.IsBinaryReferenceOperation())
                        {
                            format = "{0}{1}{2}";
                        }

                        return string.Format(format, childsL[0], childsL[1], childsL[2]);
                    }

                    if (input.IsUnion())
                    {
                        return $"({string.Join(",", childsL)})";
                    }

                    if (input.IsUnaryOperation())
                    {
                        return string.Join("", childsL);
                    }

                    throw new ArgumentException("Unknown function type.");

                case GrammarNames.Reference:
                    if (input.IsParentheses())
                    {
                        return $"({childs.First()})";
                    }

                    return string.Join("", childs);

                case GrammarNames.Prefix:
                    ret = string.Join("", childs);
                    // The exclamation mark token is not included in the parse tree, so we have to add that if it's a single file
                    if (input.ChildNodes.Count == 1 && input.ChildNodes[0].Is(GrammarNames.File))
                    {
                        ret += "!";
                    }
                    return ret;

                case GrammarNames.ArrayFormula:
                    return "{=" + childs.ElementAt(1) + "}";

                case GrammarNames.StructureReference:
                    ret = "";
                    var hastable = input.ChildNodes.Count == 2;
                    var contentsNode = hastable ? 1 : 0;
                    childsL = childs.ToList();
                    if (hastable) ret += childsL[0];

                    if (input.ChildNodes[contentsNode].Is(GrammarNames.StructureReferenceColumnOrKeyword))
                    {
                        ret += childsL[contentsNode];
                    }
                    else
                    {
                        ret += $"[{childsL[contentsNode]}]";
                    }

                    return ret;

                // Terms for which to print all child nodes concatenated
                case GrammarNames.ArrayConstant:
                case GrammarNames.DynamicDataExchange:
                case GrammarNames.FormulaWithEq:
                case GrammarNames.File:
                case GrammarNames.StructureReferenceExpression:
                    return string.Join("", childs);

                // Terms for which we print the childs comma-separated
                case GrammarNames.Arguments:
                case GrammarNames.ArrayRows:
                case GrammarNames.Union:
                    return string.Join(",", childs);

                case GrammarNames.ArrayColumns:
                    return string.Join(";", childs);

                case GrammarNames.ConstantArray:
                    return $"{{{childs.First()}}}";

                default:
                    // If it is not defined above and the number of childs is exactly one, we want to just print the first child
                    if (input.ChildNodes.Count == 1)
                    {
                        return childs.First();
                    }
                    throw new ArgumentException($"Could not print node of type '{input.Term.Name}'.\nThis probably means the excel grammar was modified without the print function being modified");
            }
        }

        /// <summary>
        /// Print a F# tree to a formula string
        /// </summary>
        public static string Print(this FSharpTransform.Formula result)
        {
            if (result.IsS) {
                var y = (FSharpTransform.Formula.S)result;
                var Cell = y.Item;

                var CCell = (FSharpTransform.SuperCell.C)Cell;
                string CellName = new Location(CCell.Item.Item1, CCell.Item.Item2).ToString();
                return CellName;
            }

            if (result.IsRange) {
                var y = (FSharpTransform.Formula.Range)result;
                var Cell1 = y.Item1;
                var Cell2 = y.Item2;

                string CellName1;
                string CellName2;

                if (Cell1.IsC) {
                    var CCell1 = (FSharpTransform.SuperCell.C)Cell1;
                    CellName1 = new Location(CCell1.Item.Item1, CCell1.Item.Item2).ToString();
                } else {
                    throw new ArgumentException("Unable to print dynamic tree");
                }

                if (Cell2.IsC) {
                    var CCell2 = (FSharpTransform.SuperCell.C)Cell2;
                    CellName2 = new Location(CCell2.Item.Item1, CCell2.Item.Item2).ToString();
                } else {
                    throw new ArgumentException("Unable to print dynamic tree");
                }

                return CellName1 + ":" + CellName2;
            }

            if (result.IsFunction) {
                var y = (FSharpTransform.Formula.Function)result;
                string FunctionName = y.Item1;

                if (FunctionName.Contains("(")) //it is a prefix function
                {
                    string Arguments = PrintArguments(y.Item2);
                    return FunctionName + Arguments + ")";
                }

                if (FunctionName == "") {
                    string Arguments = PrintArguments(y.Item2);
                    return "(" + FunctionName + Arguments + ")";
                } else //infix
                {
                    return Print(y.Item2.First()) + FunctionName + Print(y.Item2.ElementAt(1));
                }
            }

            if (result.IsConstant) {
                var y = (FSharpTransform.Formula.Constant)result;
                return y.Item;

            }

            if (result.IsArgumentList) {
                var y = (FSharpTransform.Formula.ArgumentList)result;
                return PrintArguments(y.Item);
            }

            throw new ArgumentException("Unable to print dynamic tree");

        }

        private static string PrintArguments(FSharpList<FSharpTransform.Formula> y)
        {
            string Arguments = "";
            foreach (FSharpTransform.Formula Argument in y) {
                Arguments += Print(Argument) + ",";
            }

            Arguments = RemoveFinalSymbol(Arguments);
            return Arguments;
        }

        private static string RemoveFinalSymbol(string input)
        {
            input = input.Substring(0, input.Length - 1);
            return input;
        }
    }
}
