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
    public class FSharpTransformationRule : IComparable<FSharpTransformationRule>
    {
        public string Name;
        public ParseTreeNode from;
        public ParseTreeNode to;
        public double priority;

        private readonly Lazy<FSharpTransform.Formula> fromFSharpTree;
        private readonly Lazy<FSharpTransform.Formula> toFSharpTree;

        public FSharpTransformationRule()
        {
            fromFSharpTree = new Lazy<FSharpTransform.Formula>(() => CreateFSharpTree(from));
            toFSharpTree = new Lazy<FSharpTransform.Formula>(() => CreateFSharpTree(to));
        }

        public int CompareTo(FSharpTransformationRule y)
        {
            return priority.CompareTo(y.priority);
        }

        /// <summary>
        /// Parse a BumbleBee or Excel formula to a parse tree
        /// </summary>
        /// <param name="input">Formula without =</param>
        /// <returns>Parse tree on success or null on error.</returns>
        public ParseTreeNode ParseToTree(string input)
        {
            return FSharpFormulaHelper.ParseToTree(input);
        }

        /// <summary>
        /// Change a C# parse tree to a F# parse tree
        /// </summary>
        public FSharpTransform.Formula CreateFSharpTree(ParseTreeNode input)
        {
            return FSharpFormulaHelper.CreateFSharpTree(input);
        }

        public bool CanBeAppliedonBool(string formula)
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree source = P.ParseToTree(formula);

            FSharpTransform.Formula FFrom = fromFSharpTree.Value;
            FSharpTransform.Formula FSource = CreateFSharpTree(source.Root);

            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public bool CanBeAppliedonBool(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = fromFSharpTree.Value;
            FSharpTransform.Formula FSource = CreateFSharpTree(source);
            
            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public FSharpMap<char, FSharpTransform.mapElement> CanBeAppliedonMap(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = fromFSharpTree.Value;
            FSharpTransform.Formula FSource = CreateFSharpTree(source);

            return FSharpTransform.CanBeAppliedonMap(FFrom, FSource);
        }

        /// <summary>
        /// Apply this transformation rule on a formula
        /// </summary>
        public string ApplyOn(string formula)
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree source = P.ParseToTree(formula);

            FSharpTransform.Formula FFrom = fromFSharpTree.Value;
            FSharpTransform.Formula FTo = toFSharpTree.Value;
            FSharpTransform.Formula FSource = CreateFSharpTree(source.Root);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return Print(result);
        }

        /// <summary>
        /// Apply this transformation rule on a parse tree
        /// </summary>
        public FSharpTransform.Formula ApplyOn(ParseTreeNode source)
        {
            FSharpTransform.Formula FFrom = fromFSharpTree.Value;
            FSharpTransform.Formula FTo = toFSharpTree.Value;
            FSharpTransform.Formula FSource = CreateFSharpTree(source);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return result;
        }

        /// <summary>
        /// Print a F# tree to a formula string
        /// </summary>
        public string Print(FSharpTransform.Formula result)
        {
            return FSharpFormulaHelper.Print(result);
        }

        
    }
}
