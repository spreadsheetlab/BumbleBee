using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using Irony.Parsing;
using Infotron.Util;
using Microsoft.FSharp.Collections;
using System.Diagnostics;
using FSharpEngine;
using XLParser;


namespace Infotron.FSharpFormulaTransformation
{
    public class FSharpTransformationRule : IComparable<FSharpTransformationRule>
    {
        public string Name;
        public ParseTreeNode from;
        public ParseTreeNode to;
        public double priority;

        // A lot of code does the initalization itself or with object initializers
        public FSharpTransformationRule()
        {
        }

        public FSharpTransformationRule(string Name, string from, string to, double priority = 0)
        {
            this.Name = Name;
            this.from = ParseToTree(from);
            this.to = ParseToTree(to);
            this.priority = priority;

            if (this.from == null)
            {
                throw new ArgumentException($"Could not parse \"{from}\"", nameof(from));
            }
            if (this.to == null)
            {
                throw new ArgumentException($"Could not parse \"{to}\"", nameof(to));
            }


        }

        private FSharpTransform.Formula _fromFSharpTree;
        private FSharpTransform.Formula _toFSharpTree;
        private FSharpTransform.Formula fromFSharpTree => _fromFSharpTree ?? (_fromFSharpTree = CreateFSharpTree(from));
        private FSharpTransform.Formula toFSharpTree => _toFSharpTree ?? (_toFSharpTree = CreateFSharpTree(to));

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
            return input.CreateFSharpTree();
        }

        public bool CanBeAppliedonBool(string formula)
        {
            var source = ExcelFormulaParser.Parse(formula);

            var FFrom = fromFSharpTree;
            var FSource = CreateFSharpTree(source);

            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public bool CanBeAppliedonBool(ParseTreeNode source)
        {
            var FFrom = fromFSharpTree;
            var FSource = CreateFSharpTree(source);
            
            return FSharpTransform.CanBeAppliedonBool(FFrom, FSource);
        }

        public FSharpMap<char, FSharpTransform.mapElement> CanBeAppliedonMap(ParseTreeNode source)
        {
            var FFrom = fromFSharpTree;
            var FSource = CreateFSharpTree(source);

            return FSharpTransform.CanBeAppliedonMap(FFrom, FSource);
        }

        /// <summary>
        /// Apply this transformation rule on a formula
        /// </summary>
        public string ApplyOn(string formula)
        {
            var source = ExcelFormulaParser.Parse(formula);

            var FFrom = fromFSharpTree;
            var FTo = toFSharpTree;
            var FSource = CreateFSharpTree(source);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return Print(result);
        }

        /// <summary>
        /// Apply this transformation rule on a parse tree
        /// </summary>
        public FSharpTransform.Formula ApplyOn(ParseTreeNode source)
        {
            var FFrom = fromFSharpTree;
            var FTo = toFSharpTree;
            var FSource = CreateFSharpTree(source);

            var result = FSharpTransform.ApplyOn(FTo, FFrom, FSource);

            return result;
        }

        /// <summary>
        /// Print a F# tree to a formula string
        /// </summary>
        public string Print(FSharpTransform.Formula result)
        {
            return result.Print();
        }

        
    }
}
