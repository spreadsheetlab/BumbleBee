using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelAddIn3.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using Infotron.Parsing;
using Infotron.Util;
using Irony.Parsing;

namespace ExcelAddIn3.Refactorings
{
    /// <summary>
    /// Group a set of references
    /// </summary>
    public class GroupReferences : FormulaRefactoring
    {
        private _Worksheet excel;
        public GroupReferences(_Worksheet excel)
        {
            this.excel = excel;
        }

        public GroupReferences(){}

        public override void Refactor(Range applyto)
        {
            excel = applyto.Worksheet;
            base.Refactor(applyto);
        }

        public override ParseTreeNode Refactor(ParseTreeNode applyto)
        {
            if (excel == null)
            {
                throw new InvalidOperationException("Must have reference to Excel worksheet to group references");
            }
            var targetFunctions = ExcelFormulaParser.AllNodes(applyto)
                .Where(IsTargetFunction);

            foreach (var function in targetFunctions)
            {
                var arguments = function.ChildNodes[1].ChildNodes;

                // Group ArrayAsArgument arguments
                foreach (var arg in arguments.Where(arg => arg.Is(GrammarNames.ArrayAsArgument)))
                {
                    GroupReferenceList(arg.ChildNodes[0].ChildNodes);
                }

                // If this is a varags function group all arguments
                if (varargsFunctions.Contains(ExcelFormulaParser.GetFunction(function)))
                {
                    GroupReferenceList(arguments);
                }
            }

            return applyto;
        }

        private void GroupReferenceList(ParseTreeNodeList arguments)
        {
            var togroup = arguments
                    .Where(NodeCanBeGrouped);
            var toNotGroup = arguments
                .Where(x => !NodeCanBeGrouped(x));

            var grouped = GroupTheReferences(togroup)
                .OrderBy(x => x) // Sort references alphabetically
                .Select(x => x.Parse()); // Make them parsetreenodes again

            var newargs = toNotGroup.Concat(grouped).ToList();
            arguments.Clear();
            arguments.AddRange(newargs);
        }

        public override bool CanRefactor(ParseTreeNode applyto)
        {
            return ExcelFormulaParser.AllNodes(applyto).Any(IsTargetFunction);
        }

        private static bool IsTargetFunction(ParseTreeNode node)
        {
            return
                    // Not interested in not-functions
                    ExcelFormulaParser.IsFullFunction(node)
                    // Or functions without arguments
                    && node.ChildNodes[1].ChildNodes.Any() 
                    && (varargsFunctions.Contains(ExcelFormulaParser.GetFunction(node))
                        // Functions have an arrayasargument parameter
                        || node.ChildNodes[1].ChildNodes.Any(n => n.Is(GrammarNames.ArrayAsArgument))
                       )
                   ;
        }

        private static bool NodeCanBeGrouped(ParseTreeNode node)
        {
            // can be grouped if the node is a reference
            var relevant = ExcelFormulaParser.SkipToRevelantChildNodes(node);
            return relevant.Is(GrammarNames.Reference)
                // no named ranges
                && !ExcelFormulaParser.AllNodes(node).Any(x=>x.Is(GrammarNames.NamedRange))
                // no vertical or horizontal ranges
                && !(relevant.ChildNodes[0].ChildNodes[0].Is(GrammarNames.Range) && relevant.ChildNodes[0].ChildNodes[0].ChildNodes.Count == 1);
        }

        /// <summary>
        /// Takes a list of references and return a grouped list of references
        /// </summary>
        private IEnumerable<string> GroupTheReferences(IEnumerable<ParseTreeNode> references)
        {
            var refs = references.Select(r => new {abs = checkAbsolute(r),reference = r.Print()}).ToList();
            var output = new List<string>();
            // We don't do anything with things that mix absolute and relative markers
            output.AddRange(refs.Where(x => x.abs.mixed).Select(x => x.reference));

            // Now make excel group everything, divided by absolute/relative row/columns
            var absoluteCategories = refs
                .Where(x => !x.abs.mixed)
                .GroupBy(x => new {colA = x.abs.colAbsolute, rowA = x.abs.rowAbsolute})
                ;

            foreach (var grouping in absoluteCategories)
            {
                var ranges = grouping.Select(x => excel.Range[x.reference]);
                //var rangestring = String.Join((string)excel.Application.International[XlApplicationInternational.xlListSeparator], grouping.Select(x => x.reference));
                //var union = excel.Range[rangestring];
                output.AddRange(GroupRanges(ranges).Address[grouping.Key.rowA,grouping.Key.colA].Split(','));
            }

            return output;
        }

        private static Range GroupRanges(IEnumerable<Range> ranges)
        {
            var size = ranges.Count();
            switch (size)
            {
                case 0:
                    throw new ArgumentException("Cannot group 0 ranges");
                case 1:
                    var r = ranges.First();
                    return r.Application.Union(r, r);

            }
            if (size <= 30)
            {
                var unionarguments = Enumerable.Repeat(Type.Missing, 30).ToArray();
                var i = 0;
                foreach (var r in ranges)
                {
                    unionarguments[i] = r;
                    i++;
                }
                return ((Range)unionarguments[0]).Application.Union(
                    (Range) unionarguments[0],
                    (Range) unionarguments[1],
                    unionarguments[2],
                    unionarguments[3],
                    unionarguments[4],
                    unionarguments[5],
                    unionarguments[6],
                    unionarguments[7],
                    unionarguments[8],
                    unionarguments[9],
                    unionarguments[10],
                    unionarguments[11],
                    unionarguments[12],
                    unionarguments[13],
                    unionarguments[14],
                    unionarguments[15],
                    unionarguments[16],
                    unionarguments[17],
                    unionarguments[18],
                    unionarguments[19],
                    unionarguments[20],
                    unionarguments[21],
                    unionarguments[22],
                    unionarguments[23],
                    unionarguments[24],
                    unionarguments[25],
                    unionarguments[26],
                    unionarguments[27],
                    unionarguments[28],
                    unionarguments[29]
                    );
            }

            return GroupRanges(ranges.Batch(30).Select(GroupRanges));

        }

        /// <summary>
        /// Check if all cells references have the same row/col absolute type or if it's mixed
        /// </summary>
        private static Absolute checkAbsolute(ParseTreeNode reference)
        {
            var cells = ExcelFormulaParser.AllNodes(reference).Where(x => x.Is(GrammarNames.Cell));
            bool first = true;
            var a = new Absolute();
            var locs = cells.Select(cell => new Location(cell.Print()));
            foreach (var l in locs)
            {
                if (first)
                {
                    a.colAbsolute = l.ColumnFixed;
                    a.rowAbsolute = l.RowFixed;
                    a.mixed = false;
                    first = false;
                }
                else
                {
                    if (a.colAbsolute != l.ColumnFixed || a.rowAbsolute != l.RowFixed)
                    {
                        a.mixed = true;
                    }
                }
            }
            return a;
        }

        protected override RangeShape.Flags AppliesTo { get { return RangeShape.Flags.NonEmpty; } }

        /// <summary>
        /// List of functions on which multiple arguments act the same as a single ArrayAsArgument parameter.
        /// Basically these are all the functions that have only a single ArrayAsArgument parameter
        /// </summary>
        /// <example>
        /// SUM(A1,B5:B10,K9) is identical to SUM((A1,B5:B10,K9)) and thus belongs in the list
        /// SMALL(A1,B2) is not identical to  SMALL((A1,B2)) and thus doesn't belong in it
        /// </example>
        private static readonly ISet<String> varargsFunctions = new HashSet<string>()
        {
            // Source: http://superuser.com/questions/447492/is-there-a-union-operator-in-excel
            "SUM",
            "COUNT",
            "COUNTA",
            "COUNTBLANK",
            "LARGE",
            "MIN",
            "MAX",
            "AVERAGE",
        };

        private class Absolute
        {
            public bool colAbsolute = false;
            public bool rowAbsolute = false;
            // If it is mixed, e.g. the range $A1:A$7
            public bool mixed = false;
        }
    }


    // Source: https://code.google.com/p/morelinq/source/browse/MoreLinq/Batch.cs?r=f85495b139a19bce7df2be98ad88754ba8932a28
    #region License and Terms
    // MoreLINQ - Extensions to LINQ to Objects
    // Copyright (c) 2008-2011 Jonathan Skeet. All rights reserved.
    // 
    // Licensed under the Apache License, Version 2.0 (the "License");
    // you may not use this file except in compliance with the License.
    // You may obtain a copy of the License at
    // 
    //     http://www.apache.org/licenses/LICENSE-2.0
    // 
    // Unless required by applicable law or agreed to in writing, software
    // distributed under the License is distributed on an "AS IS" BASIS,
    // WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    // See the License for the specific language governing permissions and
    // limitations under the License.
    #endregion
    public static class MoreEnumerable
    {
        /// <summary>
        /// Batches the source sequence into sized buckets.
        /// </summary>
        /// <typeparam name="TSource">Type of elements in <paramref name="source"/> sequence.</typeparam>
        /// <param name="source">The source sequence.</param>
        /// <param name="size">Size of buckets.</param>
        /// <returns>A sequence of equally sized buckets containing elements of the source collection.</returns>
        /// <remarks> This operator uses deferred execution and streams its results (buckets and bucket content).</remarks>
        public static IEnumerable<IEnumerable<TSource>> Batch<TSource>(this IEnumerable<TSource> source, int size)
        {
            return Batch(source, size, x => x);
        }

        /// <summary>
        /// Batches the source sequence into sized buckets and applies a projection to each bucket.
        /// </summary>
        /// <typeparam name="TSource">Type of elements in <paramref name="source"/> sequence.</typeparam>
        /// <typeparam name="TResult">Type of result returned by <paramref name="resultSelector"/>.</typeparam>
        /// <param name="source">The source sequence.</param>
        /// <param name="size">Size of buckets.</param>
        /// <param name="resultSelector">The projection to apply to each bucket.</param>
        /// <returns>A sequence of projections on equally sized buckets containing elements of the source collection.</returns>
        /// <remarks> This operator uses deferred execution and streams its results (buckets and bucket content).</remarks>
        public static IEnumerable<TResult> Batch<TSource, TResult>(this IEnumerable<TSource> source, int size,
            Func<IEnumerable<TSource>, TResult> resultSelector)
        {
            if(source == null) throw new ArgumentNullException("source");
            if(size < 1) throw new ArgumentException("Must be positive", "size");
            if (resultSelector == null) throw new ArgumentNullException("resultSelector");
            return BatchImpl(source, size, resultSelector);
        }

        private static IEnumerable<TResult> BatchImpl<TSource, TResult>(this IEnumerable<TSource> source, int size,
            Func<IEnumerable<TSource>, TResult> resultSelector)
        {
            Debug.Assert(source != null);
            Debug.Assert(size > 0);
            Debug.Assert(resultSelector != null);

            TSource[] bucket = null;
            var count = 0;

            foreach (var item in source)
            {
                if (bucket == null)
                {
                    bucket = new TSource[size];
                }

                bucket[count++] = item;

                // The bucket is fully buffered before it's yielded
                if (count != size)
                {
                    continue;
                }

                // Select is necessary so bucket contents are streamed too
                yield return resultSelector(bucket.Select(x => x));

                bucket = null;
                count = 0;
            }

            // Return the last bucket with all remaining elements
            if (bucket != null && count > 0)
            {
                yield return resultSelector(bucket.Take(count));
            }
        }
    }
}