using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using BumbleBee.Refactorings.Util;
using Irony.Parsing;
//using Microsoft.Office.Interop.Excel;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
using XLParser;
using Infotron.Util;

namespace BumbleBee.Refactorings
{
    class AgregrateToConditionalAggregrate : RangeRefactoring
    {
        public override void Refactor(ExcelRaw.Range applyto)
        {
            if(!CanRefactor(applyto)) throw new ArgumentException("Cannot refactor this range");

            // Refactor + to SUM first so we can focus on that case
            if (_opToAggregate.CanRefactor(applyto))
            {
                _opToAggregate.Refactor(applyto);
            }

            var node = Helper.Parse(applyto).SkipToRelevant(false);
            var fname = node.GetFunction();
            var fargs = node.GetFunctionArguments();

            PrefixInfo prefix;
            bool columnEqual;
            List<int> summedColumns;
            bool rowEqual;
            List<int> summedRows;
            if (!checkRowAndColumns(node, out prefix, out columnEqual, out summedColumns, out rowEqual, out summedRows))
            {
                throw new ArgumentException("All references cells must be in the same row or column");
            }

            summedColumns.Sort();
            summedRows.Sort();

            var summedColumnsSet = new HashSet<int>(summedColumns);
            var summedRowsSet = new HashSet<int>(summedRows);

            if (!columnEqual) throw new NotImplementedException("Cant do this for rows yet");
            ExcelRaw.Range summedRange = null;
            ExcelRaw.Worksheet worksheet = null;
            ExcelRaw.Range usedRange = null;
            ExcelRaw.Range usedColumns = null;
            ExcelRaw.Range usedRows = null;
            try
            {
                summedRange = applyto.Precedents;


                // Check if we have the correct range
                var precendentRows = summedRange.Cast<ExcelRaw.Range>().Select(cell =>
                {
                    var addr = cell.Address[false, false];
                    Marshal.ReleaseComObject(cell);
                    return new Location(addr).Row1;
                }).OrderBy(x=>x);
                Debug.Assert(summedRows.SequenceEqual(precendentRows), "Precedents given by Excel did now correspond to summed rows");

                worksheet = applyto.Worksheet;
                //cells = worksheet.Cells;
                // Find last filled column in worksheet
                usedRange = worksheet.UsedRange;
                usedColumns = usedRange.Columns;
                usedRows = usedRange.Rows;
                var usedRowsCount = usedRows.Count;

                // Determiner column: column we can use as the subject for the SUMIF predicate
                var determiners = usedColumns.Cast<ExcelRaw.Range>().Select(column => {
                    ExcelRaw.Range columncells = null;
                    ExcelRaw.Range firstCell = null;
                    try
                    {
                        columncells = column.Cells;
                        firstCell = columncells[summedRows[0], 1];
                        object candidatevalue = firstCell.Value2;
                        
                        // Check if [Column,Firstrow] contains a value
                        if (candidatevalue == null) return null;

                        // Check if all summed rows contain the same value
                        // Check if none of the other rows contain that value
                        if (Enumerable.Range(1, usedRowsCount).All(row =>
                        {
                            ExcelRaw.Range cell = columncells[row, 1];
                            object v = cell.Value2;
                            cell.ReleaseCom();
                            // Empty cell, candidate can't be empty so we can ignore this
                            if (v == null) return true;
                            bool isSummed = summedRowsSet.Contains(row);
                            // All summed rows must be candidatevalue
                            // All not-summed rows must not be candidatevalue
                            return (isSummed && candidatevalue.Equals(v)) || (!isSummed && !candidatevalue.Equals(v));
                        }))
                        {
                            // We found a candidatecolumn!
                            return Tuple.Create(candidatevalue, column.Column);
                        }

                        return null;
                    }
                    finally
                    {
                        firstCell.ReleaseCom();
                        columncells.ReleaseCom();
                        column.ReleaseCom();
                    }
                });
                var determiner = determiners.FirstOrDefault(found => found != null);

                // If we didn't find a candidate, do nothing
                if (determiner == null) return;
                string determinerValue = (determiner.Item1 is double) ? determiner.Item1.ToString() : $"\"{determiner.Item1}\"";
                var determinerColumn = AuxConverter.ConvertColumnToStr(determiner.Item2-1);
                var summedColumn = AuxConverter.ConvertColumnToStr(summedColumns[0]-1);
                
                var formula = $"{fname}IF({determinerColumn}:{determinerColumn},{determinerValue},{summedColumn}:{summedColumn})";
                try
                {
                    applyto.Formula = $"={formula}";
                }
                catch (COMException)
                {
                    throw new InvalidOperationException($"Refactoring created invalid formula <<{formula}>>");
                }
            }
            finally
            {
                usedRange.ReleaseCom();
                usedColumns.ReleaseCom();
                usedRows.ReleaseCom();
                worksheet.ReleaseCom();
                summedRange.ReleaseCom();
            }

        }

        private static bool IsSingleCellReference(ParseTreeNode arg)
        {
            arg = arg.SkipToRelevant(false);
            // Cell can have prefix
            return arg.ChildNodes[(arg.ChildNodes.Count == 1 ? 0 : 1)].Is(GrammarNames.Cell);
        }

        private readonly OpToAggregate _opToAggregate = new OpToAggregate();

        public override bool CanRefactor(ExcelRaw.Range applyto)
        {
            // Shape check
            if (!base.CanRefactor(applyto)) return false;

            var node = Helper.Parse(applyto);

            // Refactor op to aggregate first so we only have to handle that case
            if (_opToAggregate.CanRefactor(node))
            {
                node = _opToAggregate.Refactor(node);
            }

            node = node.SkipToRelevant(false);
            if (!node.IsNamedFunction()) return false;
            var funcname = node.GetFunction();
            if(!functions.Contains(funcname)) return false;

            PrefixInfo prefix;
            bool columnEqual;
            List<int> columns;
            bool rowEqual;
            List<int> rows;
            if (!checkRowAndColumns(node, out prefix, out columnEqual, out columns, out rowEqual, out rows)) return false;
            
            // If we only have a single precedent, no sense in summing
            if (columnEqual && rowEqual) return false;

            // If we encounter a row multiple times, we can't do the refactoring
            if ((columnEqual && rows.Count != rows.Distinct().Count())
              ||(rowEqual && columns.Count != columns.Distinct().Count()))
            {
                return false;
            }

            return true;
        }

        private static bool checkRowAndColumns(ParseTreeNode node, out PrefixInfo prefix, out bool columnEqual, out List<int> columns, out bool rowEqual, out List<int> rows)
        {
            var fargs = node.GetFunctionArguments().Select(arg => arg.SkipToRelevant()).ToList();

            prefix = null;
            columnEqual = true;
            columns = new List<int>();
            rowEqual = true;
            rows = new List<int>();

            // Check if all arguments are single-cell references
            // And if all ar in the same column/row and prefix

            // Check first cell for initial values to compare
            if (!fargs[0].Is(GrammarNames.Reference)) return false;
            if (!(fargs[0].ChildNodes[fargs[0].ChildNodes.Count == 1 ? 0 : 1]).Is(GrammarNames.Cell)) return false;
            prefix = (fargs[0].ChildNodes.Count == 1 ? null : fargs[0].ChildNodes[0])?.GetPrefixInfo();
            var loc = new Location((fargs[0].ChildNodes[fargs[0].ChildNodes.Count == 1 ? 0 : 1]).Print());

            var column = loc.Column1;
            var row = loc.Row1;

            foreach (var refnode in fargs)
            {
                if (!refnode.Is(GrammarNames.Reference)) return false;
                var index = refnode.ChildNodes.Count == 1 ? 0 : 1;
                // Check if it is a cell
                if (!refnode.ChildNodes[index].Is(GrammarNames.Cell)) return false;
                // Check if all prefixes are equal
                if (index == 0 && prefix != null) return false;
                if (index == 1 && !refnode.ChildNodes[0].GetPrefixInfo().Equals(prefix)) return false;
                loc = new Location(refnode.ChildNodes[index].Print());
                // Add rows/columns to the list and check if they are equal
                if (columnEqual && column != loc.Column1) columnEqual = false;
                if (rowEqual && row != loc.Row1) rowEqual = false;
                if (!columnEqual && !rowEqual) return false;
                rows.Add(loc.Row1);
                columns.Add(loc.Column1);
            }
            return true;
        }

        private static readonly IReadOnlyList<string> functions = new [] { "AVERAGE", "COUNT","SUM" };

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.SingleCell;

        /// <summary>Tests whether this </summary>
        private static bool isArgumentArray(ParseTreeNode n)
        {
            if (!n.Is(GrammarNames.Arguments))
            {
                return false;
            }
            return false;
        }
    }
}
