using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Infotron.FSharpFormulaTransformation;
using Infotron.PerfectXL.SmellAnalyzer;
using System.ComponentModel;
using System.Diagnostics;
using GemBox.Spreadsheet;
using PerfectXL.Domain.Observation;
using Infotron.PerfectXL.SmellAnalyzer.SmellAnalyzer;
using System.Drawing;
using Infotron.PerfectXL.DataModel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ExcelAddIn3.Refactorings;
using ExcelAddIn3.TaskPanes;
using Infotron.Parsing;
using FSharpEngine;
using Irony.Parsing;
using Microsoft.Office.Tools;

namespace ExcelAddIn3
{
    public enum ApplyTo
    {
        Range,
        Worksheet,
        Workbook
    }

    public class HighlightedCell
    {
        public Range Cell;
        public Object OriginalPattern;
        public Object OriginalColor;
        public Object OriginalComment;

        public HighlightedCell(Range cell,
            Object originalPattern,
            Object originalColor,
            Object originalComment)
        {
            this.Cell = cell;
            this.OriginalPattern = originalPattern;
            this.OriginalColor = originalColor;
            this.OriginalComment = (originalComment != null) ? ((Comment)originalComment).Text() : null;
        }

        public void Reset(){
            Cell.Interior.Color = OriginalColor;
            Cell.Interior.Pattern = OriginalPattern;
            if (Cell.Comment != null) Cell.Comment.Delete();
            if (OriginalComment != null) Cell.AddComment(OriginalComment.ToString());
        }

        public void Apply(Smell smell)
        {
            Cell.Interior.Pattern = XlPattern.xlPatternSolid;
            Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);

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

        public override bool Equals(System.Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            var smellyCell = obj as HighlightedCell;
            if (smellyCell == null)
            {
                return false;
            }

            return (Cell.Address == smellyCell.Cell.Address);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    public partial class BBAddIn
    {
        /// Enable to profile speed where implemented
        internal const bool PROFILE = true;

        public Ribbon1 theRibbon;
        readonly List<FSharpTransformationRule> AllTransformations = new List<FSharpTransformationRule>();
        public AnalysisController AnalysisController;
        private readonly ISet<HighlightedCell> smellyCells = new HashSet<HighlightedCell>();
        private readonly ISet<HighlightedCell> transformationCells = new HashSet<HighlightedCell>();

        private static string RemoveFirstSymbol(string input)
        {
            return input.Substring(1);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            theRibbon = new Ribbon1();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { theRibbon });
            
        }

        public void AddSheetBumbleBeeTransformations()
        {
            Excel.Worksheet selectedSheet = Application.ActiveSheet;

            var workbook = Application.ActiveWorkbook;
            Excel.Worksheet BumbleBeeSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            BumbleBeeSheet.Name = "_bumblebeerules";
            loadExampleTransformations(BumbleBeeSheet);
            selectedSheet.Select();

            InitializeBB();
        }

        public void InitializeBB()
        {
            //initialize transformations
            Excel.Worksheet Sheet = GetWorksheetByName("_bumblebeerules");
            if (Sheet == null)
            {
                theRibbon.groupInitialize.Visible = true;
                theRibbon.groupBumbleBee.Visible = false;
                return;
            }

            //initialize smell controls
            theRibbon.selectSmellType.Items.Clear();
            theRibbon.selectSmellType.Enabled = false;
            theRibbon.groupInitialize.Visible = false;
            theRibbon.groupBumbleBee.Visible = true;

            //find last filled cells
            int Lower = 50;

            AllTransformations.Clear();

            for (int i = 1; i <= Lower; i++)
            {
                string From = ((Range)Sheet.Cells.Item[i, 1]).Value;
                if (From != null)
                {
                    string To = ((Range)Sheet.Cells.Item[i, 2]).Value;
                    double prio = ((Range)Sheet.Cells.Item[i, 3]).Value;
                    string Name = ((Range)Sheet.Cells.Item[i, 4]).Value;

                    FSharpTransformationRule S = new FSharpTransformationRule();
                    S.from = S.ParseToTree(From);
                    S.to = S.ParseToTree(To);
                    S.priority = prio;
                    S.Name = Name;

                    AllTransformations.Add(S);

                }
            }


            //order by priority
            AllTransformations.Sort();          
        }

        private void InitializeTransformations()
        {
            theRibbon.Preview.Text = "";
            theRibbon.valuePreview.Text = "";
            theRibbon.valuePreview.ShowImage = false;
            theRibbon.dropDown1.Items.Clear();
            decolorCells(transformationCells);
        }

        public void FindApplicableTransformations()
        {
            Log("FindApplicableTransformations");

            InitializeTransformations();
            //Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Range selectedRange = ((Range)Application.Selection);
            Range selectedCell = (Range)selectedRange.Item[1, 1];
            string Formula = selectedCell.Formula;

            if (selectedCell.HasFormula && Formula.Length > 0)
            {
                Formula = RemoveFirstSymbol(Formula);

                foreach (FSharpTransformationRule t in AllTransformations)
                {
                    if (t.CanBeAppliedonBool(Formula))
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = t.Name;
                        theRibbon.dropDown1.Items.Add(item);
                    }
                }
                if (theRibbon.dropDown1.Items.Count > 0)
                {
                    MakePreview();
                }
                else
                {
                    MessageBox.Show("No applicable rewrites found.");
                }
            }
            else
            {
                MessageBox.Show("Cell does not contain a formula.");
            }
        }

        private void Log(string LogMessage)
        {
            string currentWorkbookFilePath = this.Application.ActiveWorkbook.Path;
            string LogFileName = "spreadsheets.log";
            string LogFile = System.IO.Path.Combine(currentWorkbookFilePath, LogFileName);
            var file = new System.IO.StreamWriter(LogFile, true);
            file.WriteLine(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ", " + LogMessage);
            file.Close();
        }

        public void MakePreview()
        {
            decolorCells(transformationCells);
            if (theRibbon.dropDown1.Items.Count > 0) //if we have transformations available
            {
                FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

                Excel.Range R = ((Excel.Range)Application.Selection);
                string formula = RemoveFirstSymbol(R.Item[1, 1].Formula);
                theRibbon.Preview.Text = T.ApplyOn(formula);
                theRibbon.valuePreview.Text = getValue(R.Item[1, 1], theRibbon.Preview.Text);
                theRibbon.valuePreview.ShowImage = (theRibbon.valuePreview.Text != R.Item[1, 1].Value.ToString());

                if (R.Count == 1)
                {
                    foreach (Excel.Worksheet worksheet in Application.Worksheets)
                    {
                        applyInRange(T, worksheet.UsedRange, true);
                    }
                }
                else
                {
                    applyInRange(T, Application.Selection, true);
                }
            }
        }

        private String getValue(Range cell, String formula)
        {
            string value;
            string currentFormula = cell.Formula;
            cell.Formula = "=" + formula;
            value = cell.Value.ToString();
            cell.Formula = currentFormula;
            return value;
        }

        private bool valueChanges(Range cell, String formula)
        {
            return getValue(cell, formula) != cell.Value.ToString();
        }

        public void ApplyTransformation(ApplyTo scope)
        {
            if (theRibbon.dropDown1.SelectedItem == null)
            {
                Log("ApplyTransformation tried with empty dropdown");
                return;
            }

            decolorCells(transformationCells);

            Log("Apply in " + scope.ToString() + " transformation " + theRibbon.dropDown1.SelectedItem.Label);

            FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

            switch (scope)
            {
                case ApplyTo.Range:
                    applyInRange(T, Application.Selection);
                    break;
                case ApplyTo.Worksheet:
                    applyInRange(T, Application.ActiveSheet.UsedRange);
                    break;
                case ApplyTo.Workbook:
                    foreach (Excel.Worksheet worksheet in Application.Worksheets)
                    {
                        applyInRange(T, worksheet.UsedRange);
                    }
                    break;
            }

            //after applying, we want to empty the preview box, find new rewrites and show them (in dropdown and preview)
            FindApplicableTransformations();
            MakePreview();
        }

        private void applyInRange(FSharpTransformationRule T, Excel.Range Range, Boolean previewOnly = false)
        {
            foreach (Excel.Range cell in Range.Cells)
            {
                if (cell.HasFormula)
                {
                    var Formula = RemoveFirstSymbol(cell.Formula);
                    if (T.CanBeAppliedonBool(Formula))
                    {
                        if (previewOnly)
                        {
                            var transformationCell = new HighlightedCell(cell, cell.Interior.Pattern, cell.Interior.Color, cell.Comment);
                            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            if (!transformationCells.Any(x => x.Equals(transformationCell)))
                            {
                                transformationCells.Add(transformationCell);
                            }
                        }
                        else
                        {
                            var transformedFormula = T.ApplyOn(Formula);
                            if (valueChanges(cell, transformedFormula))
                            {
                                if(MessageBox.Show("The transformation causes the value of cell " +
                                    cell.Worksheet.Name + ":" + cell.get_Address(false,false,Excel.XlReferenceStyle.xlA1) +
                                    " to change from " + cell.Value + " to " + getValue(cell, transformedFormula) +
                                    ". Do you want to continue?",
                                    "Alert: Cell value change",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Warning)
                                    == DialogResult.No)
                                    return;
                            }
                            cell.Formula = "=" + transformedFormula;
                            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        }
                    }
                }
            }
        }

        public void ColorSmells()
        {
            InitializeTransformations();
            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");

            if (!Application.ActiveWorkbook.Saved)
            {
                Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }

            if (!Application.ActiveWorkbook.Saved) {
                MessageBox.Show("The workbook must be saved before analysis. Aborting.");
                return;
            }

            AnalysisController = new AnalysisController
            {
                Worker = new BackgroundWorker { WorkerReportsProgress = true },
                // Seems like option has been removed by fecf71ad4d72daf5ad7f843a95ee00e07de6a25b and doesn't seem to have a replacement, maybe Preprocessors?
                //AnalysisMaxRows = 10000,
                Filename = Application.ActiveWorkbook.FullName
            };

            // Createriskmaps option was removed by 025a29a1b845d41850a0e4fd3ae2271d62933e55 and no direct replacement in same commmit
            AnalysisController.RunAnalysis();

            if (!AnalysisController.Spreadsheet.AnalysisSucceeded)
            {
                throw new Exception(AnalysisController.Spreadsheet.ErrorMessage);
            }

            ColorSmellsOfType("");

            LoadSmellTypesSelect();
        }

        public void SelectSmellsOfType()
        {
            InitializeTransformations();
            ColorSmellsOfType(theRibbon.selectSmellType.SelectedItem.Tag);
        }

        private void ColorSmellsOfType(String type)
        {
            decolorCells(smellyCells);

            List<Smell> smellsOfType;

            if (type == "")
            {
                smellsOfType = AnalysisController.DetectedSmells;
            }
            else
            {
                smellsOfType = AnalysisController.DetectedSmells.Where(x => x.AnalysisType.ToString() == type).ToList();
            }

            foreach (var smell in smellsOfType)
            {
                var analyzerExtension = new tmpAnalyzerExtension(smell.AnalysisType);
                if (analyzerExtension.GetMetricScore(smell.RiskValue) > MetricScore.None) ColorCell(smell);
            }
        }

        private void decolorCells(ISet<HighlightedCell> cells)
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

                var excelCell = Application.Sheets[cell.Worksheet.Name].Cells[cell.Location.Row + 1, cell.Location.Column + 1];

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
            theRibbon.selectSmellType.Items.Clear();

            foreach (var smellType in AnalysisController.DetectedSmells.Select(x => x.AnalysisType).Distinct())
            {
                tmpAnalyzerExtension analyzerExtension = new tmpAnalyzerExtension(smellType);
                if (AnalysisController.DetectedSmells.Any(x => analyzerExtension.GetMetricScore(x.RiskValue) > MetricScore.None))
                    addSelectSmellTypeItem(smellType.ToString(), analyzerExtension.SmellName);
            }

            addSelectSmellTypeItem("", "(all)", true);

            if(theRibbon.selectSmellType.Items.Count > 1) theRibbon.selectSmellType.Enabled = true;
        }

        private void addSelectSmellTypeItem(string id, string option, bool selected = false)
        {
            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = option;
            item.Tag = id;
            theRibbon.selectSmellType.Items.Add(item);
            if(selected) theRibbon.selectSmellType.SelectedItem = item;
        }

        private Excel.Worksheet GetWorksheetByName(string name)
        {
            return Application.Worksheets
                .Cast<Excel.Worksheet>()
                .FirstOrDefault(worksheet => worksheet.Name == name);
        }

        private void loadExampleTransformations(Excel.Worksheet BumbleBeeSheet)
        {
            String[,] exampleTransformations = {
                                               {"'IF([c]<[d],[c],[d])", "'MIN([c],[d])", "3", "IF to MIN"},
                                               {"'IF([c]>[d],[c],[d])", "'MAX([c],[d])", "3", "IF to MAX"},
                                               {"SUM({r})/COUNT({r})", "AVERAGE({r})", "2", "SUM and COUNT to AVERAGE"},
                                               {"[c]+[d]", "SUM([c],[d])", "4", "Plus to SUM"},
                                               {"SUM([x],SUM([y]))", "SUM([x],[y])", "5", "Remove Double SUM"},
                                               {"SUM({i,j}, {i,j+1}, [k])", "SUM({i,j}:{i,j+1},[k])", "6", "Merge Adjacent SUMs"},
                                               {"SUM({x,y}: {i,j}, {i,j+1},[k])", "SUM({x,y}:{i,j+1},[k])", "7", "Merge Adjacent SUMs1"},
                                               {"SUM({x,y}: {i,j}, {i,j+1} )", "SUM({x,y}:{i,j+1})", "8", "Merge Adjacent SUMs2"},
                                               {"([c])", "[c]", "9", "Remove Braces"}
                                               };
            for (var i = 0; i < 9; i++)
            {
                for (var j = 0; j < 4; j++)
                {
                    BumbleBeeSheet.Cells[i + 1, j + 1] = exampleTransformations[i, j];
                }
            }
        }

        // TODO: Better place / dynamic location, preferably inside source control
        private readonly string[] BumbleBeeDebugStartupfiles =
        {
            @"C:\bumblebee\startup.xlsx",
            @"C:\bumblebee\startup.xlsm"
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
#if DEBUG
            foreach (var startupfile in BumbleBeeDebugStartupfiles.Where(System.IO.File.Exists))
            {
                Application.Workbooks.Open(startupfile);
            }
#endif

            extractFormulaTp = new TaksPaneWPFContainer<ExtractFormulaTaskPane>(new ExtractFormulaTaskPane(Application));
            extractFormulaCtp = CustomTaskPanes.Add(extractFormulaTp, "Extract formula");

            RefactoringContextMenuInitialize();
            this.Application.SheetBeforeRightClick +=
                new AppEvents_SheetBeforeRightClickEventHandler(RefactorMenuEnableRelevant);

        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            InitializeBB();
        }

        private TaksPaneWPFContainer<ExtractFormulaTaskPane> extractFormulaTp;
        private CustomTaskPane extractFormulaCtp;

        public void extractFormula()
        {
            extractFormulaTp.Child.init(Application.Selection);
            extractFormulaCtp.Visible = true;
        }

        public void inlineFormula()
        {
            Range selected = Application.Selection;
            if (selected == null || selected.Count == 0)
            {
                MessageBox.Show("Select cell(s) to inline");
                return;
            }
            try
            {
                new InlineFormula().Refactor(Application.Selection);
                MessageBox.Show("Succesfully inlined selected cells");
            }
            catch (AggregateException e)
            {
                MessageBox.Show(
                    String.Format(
                        "Not all cells could be succefully inlined.\n{0}",
                        String.Join("\n\n",e.InnerExceptions.Select(ie => ie.Message))
                    )
               );
            }
            catch (Exception e)
            {
                MessageBox.Show("Unknown error");
                throw;
            }
        }

        private class RefactorMenuItem
        {
            public string MenuText { get; set; }
            public IRangeRefactoring Refactoring { get; set; }
            public bool NewGroup { get; set; }
            public Office.CommandBarButton Button { get; set; }
        }

        private readonly List<RefactorMenuItem> contextMenuRefactorings = new List<RefactorMenuItem>
        {
            new RefactorMenuItem {MenuText="+ to SUM",Refactoring=new OpToAggregate()},
            new RefactorMenuItem {MenuText="SUM to SUMIF", Refactoring=new GroupReferences()},
            new RefactorMenuItem {MenuText="Group References", Refactoring=new GroupReferences(), NewGroup = true},
            new RefactorMenuItem {MenuText="Inline Formula", Refactoring=new InlineFormula(), NewGroup = true},
            new RefactorMenuItem {MenuText="Extract Formula", Refactoring=new NotImplementedRefactoring()},
        };

        private void RefactoringContextMenuInitialize()
        {
            const string tag = "REFACTORMENU";
            var cellcontextmenu = this.Application.CommandBars["Cell"];

            // Check if menu already defined
            if (cellcontextmenu.FindControl(Office.MsoControlType.msoControlPopup, 0, tag) != null) return;

            var menu = (Office.CommandBarPopup)cellcontextmenu.Controls.Add(
                Type:Office.MsoControlType.msoControlPopup,
                Before: cellcontextmenu.Controls.Count,
                Temporary: true);
            menu.Caption = "Refactor";
            menu.BeginGroup = true;
            menu.Tag = tag;

            foreach (var menuitem in contextMenuRefactorings)
            {
                var control = (Office.CommandBarButton) menu.Controls.Add(Type: Office.MsoControlType.msoControlButton, Temporary:true);
                control.Caption = menuitem.MenuText;
                control.BeginGroup = menuitem.NewGroup;
                // Disable by default, only enable when relevant
                control.Enabled = false;
                menuitem.Button = control;
                // Create a copy of the iterator item because we're going to use it within a closure
                var refactoring = menuitem.Refactoring;
                control.Click += (Office.CommandBarButton ctrl, ref bool cancelDefault) =>
                {
                    try
                    {
                        refactoring.Refactor(Application.Selection);
                    }
                    catch (AggregateException e)
                    {
                        MessageBox.Show(
                            String.Format(
                                "Errors:\n{0}",
                                String.Join("\n\n", e.InnerExceptions.Select(ie => ie.Message))
                            )
                       );
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(String.Format("Error: {0}", e.Message));
                    }
                };
            }
        }

        private void RefactorMenuEnableRelevant(object Sh, Range Target, ref bool Cancel)
        {
            foreach (var item in contextMenuRefactorings)
            {
                Stopwatch sw;
                if (PROFILE)
                {
                    sw = Stopwatch.StartNew();
                }
                item.Button.Enabled = item.Refactoring.CanRefactor(Target);
                if (PROFILE)
                {
                    sw.Stop();
                    var cap = item.MenuText;
                    item.Button.Caption = String.Format("{0} ({1}s)", cap, sw.Elapsed.TotalSeconds);
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {    
            this.Startup += new EventHandler(ThisAddIn_Startup);
            Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
        }

        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            InitializeTransformations();
        }

        void Application_WorkbookAfterSave(Microsoft.Office.Interop.Excel.Workbook w, bool success)
        {
            InitializeBB();
        }








        
        #endregion
    }
}
