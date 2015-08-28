using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using Infotron.PerfectXL.DataModel;
using Infotron.PerfectXL.SmellAnalyzer;
using Infotron.PerfectXL.SmellAnalyzer.SmellAnalyzer;
using Microsoft.Office.Tools.Ribbon;
using PerfectXL.Domain.Observation;
using Microsoft.Office.Interop.Excel;

namespace BumbleBee
{
    public class BBColorSmells
    {
        private BBAddIn addIn;
        public AnalysisController AnalysisController;
        private readonly ISet<HighlightedCell> smellyCells = new HashSet<HighlightedCell>();

        public BBColorSmells(BBAddIn addIn)
        {
            this.addIn = addIn;
        }

        private AnalysisController controller;

        public void ColorSmells()
        {
            addIn.bbTransformations.clearTransformationsRibbon(addIn);
            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");

            if (!addIn.Application.ActiveWorkbook.Saved)
            {
                addIn.Application.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }

            if (!addIn.Application.ActiveWorkbook.Saved) {
                MessageBox.Show("The workbook must be saved before analysis. Aborting.");
                return;
            }

            controller = new AnalysisController
            {
                Worker = new BackgroundWorker { WorkerReportsProgress = true },
                // Seems like option has been removed by fecf71ad4d72daf5ad7f843a95ee00e07de6a25b and doesn't seem to have a replacement, maybe Preprocessors?
                //AnalysisMaxRows = 10000,
                Filename = addIn.Application.ActiveWorkbook.FullName
            };

            // Createriskmaps option was removed by 025a29a1b845d41850a0e4fd3ae2271d62933e55 and no direct replacement in same commmit
            controller.RunAnalysis();

            if (!controller.Spreadsheet.AnalysisSucceeded)
            {
                throw new Exception(controller.Spreadsheet.ErrorMessage);
            }

            ColorSmellsOfType("");

            LoadSmellTypesSelect();
        }

        public void SelectSmellsOfType()
        {
            addIn.bbTransformations.clearTransformationsRibbon(addIn);
            ColorSmellsOfType(addIn.theRibbon.selectSmellType.SelectedItem.Tag);
        }

        private void ColorSmellsOfType(string type)
        {
            decolorCells(smellyCells);

            List<Smell> smellsOfType = type == "" ? controller.DetectedSmells : controller.DetectedSmells.Where(x => x.AnalysisType.ToString() == type).ToList();

            foreach (var smell in smellsOfType)
            {
                var analyzerExtension = new tmpAnalyzerExtension(smell.AnalysisType);
                if (analyzerExtension.GetMetricScore(smell.RiskValue) > MetricScore.None) ColorCell(smell);
            }
        }

        public void decolorCells(ISet<HighlightedCell> cells)
        {
            foreach (HighlightedCell cell in cells)
            {
                cell.Reset();
            }
            cells.Clear();
        }

        private void ColorCell(Smell smell)
        {
            if (!smell.IsCellBased()) return;

            try
            {
                var cell = (smell.SourceType == RiskSourceType.SiblingClass) ? ((SiblingClass)smell.Source).Cells[0] : (Cell)smell.Source;

                var excelCell = addIn.Application.Sheets[cell.Worksheet.Name].Cells[cell.Location.Row + 1, cell.Location.Column + 1];

                var smellyCell = new HighlightedCell(excelCell, excelCell.Interior.Pattern, excelCell.Interior.Color, excelCell.Comment);
                smellyCell.Apply(smell);
                if (!smellyCells.Any(x => x.Equals(smellyCell)))
                {
                    smellyCells.Add(smellyCell);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        public void LoadSmellTypesSelect()
        {
            addIn.theRibbon.selectSmellType.Items.Clear();

            foreach (var smellType in controller.DetectedSmells.Select(x => x.AnalysisType).Distinct())
            {
                tmpAnalyzerExtension analyzerExtension = new tmpAnalyzerExtension(smellType);
                if (controller.DetectedSmells.Any(x => analyzerExtension.GetMetricScore(x.RiskValue) > MetricScore.None))
                    addSelectSmellTypeItem(smellType.ToString(), analyzerExtension.SmellName);
            }

            addSelectSmellTypeItem("", "(all)", true);

            if(addIn.theRibbon.selectSmellType.Items.Count > 1) addIn.theRibbon.selectSmellType.Enabled = true;
        }

        private void addSelectSmellTypeItem(string id, string option, bool selected = false)
        {
            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = option;
            item.Tag = id;
            addIn.theRibbon.selectSmellType.Items.Add(item);
            if(selected) addIn.theRibbon.selectSmellType.SelectedItem = item;
        }
    }

    public class HighlightedCell
    {
        public Range Cell;
        public object OriginalPattern;
        public object OriginalColor;
        public object OriginalComment;

        public HighlightedCell(Range cell,
            object originalPattern,
            object originalColor,
            object originalComment = null)
        {
            Cell = cell;
            OriginalPattern = originalPattern;
            OriginalColor = originalColor;
            OriginalComment = (originalComment as Comment)?.Text();
        }

        public void Reset()
        {
            Cell.Interior.Color = OriginalColor;
            Cell.Interior.Pattern = OriginalPattern;
            if (Cell.Comment != null) Cell.Comment.Delete();
            if (OriginalComment != null) Cell.AddComment(OriginalComment.ToString());
        }

        public void Apply(Smell smell)
        {
            Cell.Interior.Pattern = XlPattern.xlPatternSolid;
            Cell.Interior.Color = ColorTranslator.ToOle(Color.Red);

            var existingComment = "";
            var analyzerExtension = new tmpAnalyzerExtension(smell.AnalysisType);
            var comments = analyzerExtension.GetSmellMessage(smell);
            if (!string.IsNullOrEmpty(comments))
            {
                if (Cell.Comment != null)
                {
                    existingComment = Cell.Comment.Text() + "\n";
                    Cell.Comment.Delete();
                }
                Cell.AddComment(existingComment + comments);
                Cell.Comment.Visible = true;
            }
        }

        public override bool Equals(object obj)
        {
            var smellyCell = obj as HighlightedCell;
            if (smellyCell == null)
            {
                return false;
            }

            return (Cell.Address == smellyCell.Cell.Address);
        }

        public override int GetHashCode()
        {
            return Cell.Address.GetHashCode();
        }
    }
}