using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Infotron.FSharpFormulaTransformation;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace BumbleBee
{
    public class BBTransformations
    {
        private BBAddIn addIn;
        private List<FSharpTransformationRule> AllTransformations { get; } = new List<FSharpTransformationRule>();
        private ISet<HighlightedCell> transformationCells { get; } = new HashSet<HighlightedCell>();

        public BBTransformations(BBAddIn addIn)
        {
            this.addIn = addIn;
        }

        private static string RemoveFirstSymbol(string input)
        {
            return input.Substring(1);
        }

        public void AddSheetBumbleBeeTransformations()
        {
            Microsoft.Office.Interop.Excel.Worksheet selectedSheet = addIn.Application.ActiveSheet;

            var workbook = addIn.Application.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet BumbleBeeSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            BumbleBeeSheet.Name = "_bumblebeerules";
            loadExampleTransformations(BumbleBeeSheet);
            selectedSheet.Select();

            addIn.InitializeBB();
        }

        public void startsTransformationRules()
        {
            //initialize transformations
            ExcelRaw.Worksheet Sheet = addIn.GetWorksheetByName("_bumblebeerules");
            if (Sheet == null)
            {
                addIn.theRibbon.groupInitialize.Visible = true;
                addIn.theRibbon.groupBumbleBee.Visible = false;
                return;
            }

            //initialize smell controls
            addIn.theRibbon.selectSmellType.Items.Clear();
            addIn.theRibbon.selectSmellType.Enabled = false;
            addIn.theRibbon.groupInitialize.Visible = false;
            addIn.theRibbon.groupBumbleBee.Visible = true;

            readTransformationRules(Sheet);
        }

        private void readTransformationRules(ExcelRaw.Worksheet rules)
        {
            //find last filled cells
            int Lower = 50;

            AllTransformations.Clear();

            for (int i = 1; i <= Lower; i++)
            {
                string From = (string)rules.Cells[i, 1].Value;
                if (From != null)
                {
                    string To = (string)rules.Cells[i, 2].Value;
                    double prio = (double)rules.Cells[i, 3].Value;
                    string Name = (string)rules.Cells[i, 4].Value;

                    AllTransformations.Add(new FSharpTransformationRule(Name, From, To, prio));

                }
            }


            //order by priority
            AllTransformations.Sort();
        }

        public void FindApplicableTransformations()
        {
            addIn.Log("FindApplicableTransformations");

            addIn.bbTransformations.clearTransformationsRibbon(addIn);
            //Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range selectedRange = addIn.Application.Selection;
            Excel.Range selectedCell = selectedRange[1, 1];
            string Formula = (string)selectedCell.Formula;

            if ((bool)selectedCell.HasFormula && Formula.Length > 0)
            {
                Formula = RemoveFirstSymbol(Formula);

                foreach (FSharpTransformationRule t in AllTransformations)
                {
                    if (t.CanBeAppliedonBool(Formula))
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = t.Name;
                        addIn.theRibbon.dropDown1.Items.Add(item);
                    }
                }
                if (addIn.theRibbon.dropDown1.Items.Count > 0)
                {
                    addIn.bbTransformations.MakePreview(addIn);
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

        private string getValue(ExcelRaw.Range cell, string formula)
        {
            string value;
            string currentFormula = (string)cell.Formula;
            cell.Formula = "=" + formula;
            value = cell.Value.ToString();
            cell.Formula = currentFormula;
            return value;
        }

        private bool valueChanges(ExcelRaw.Range cell, string formula)
        {
            return getValue(cell, formula) != cell.Value.ToString();
        }

        public void ApplyTransformation(ApplyTo scope)
        {
            if (addIn.theRibbon.dropDown1.SelectedItem == null)
            {
                addIn.Log("ApplyTransformation tried with empty dropdown");
                return;
            }

            addIn.bbColorSmells.decolorCells(transformationCells);

            addIn.Log("Apply in " + scope.ToString() + " transformation " + addIn.theRibbon.dropDown1.SelectedItem.Label);

            FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == addIn.theRibbon.dropDown1.SelectedItem.Label);

            switch (scope)
            {
                case ApplyTo.Range:
                    applyInRange(T, addIn.Application.Selection);
                    break;
                case ApplyTo.Worksheet:
                    applyInRange(T, addIn.Application.ActiveSheet.UsedRange);
                    break;
                case ApplyTo.Workbook:
                    foreach (ExcelRaw.Worksheet worksheet in addIn.Application.Worksheets)
                    {
                        applyInRange(T, worksheet.UsedRange);
                    }
                    break;
            }

            //after applying, we want to empty the preview box, find new rewrites and show them (in dropdown and preview)
            FindApplicableTransformations();
            addIn.bbTransformations.MakePreview(addIn);
        }

        private void applyInRange(FSharpTransformationRule T, ExcelRaw.Range Range, bool previewOnly = false)
        {
            foreach (ExcelRaw.Range cell in Range.Cells)
            {
                if ((bool)cell.HasFormula)
                {
                    var Formula = RemoveFirstSymbol((string)cell.Formula);
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
                                                   cell.Worksheet.Name + ":" + cell.get_Address(false,false,Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1) +
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

        private void loadExampleTransformations(Microsoft.Office.Interop.Excel.Worksheet BumbleBeeSheet)
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

        public void MakePreview(BBAddIn addIn)
        {
            addIn.bbColorSmells.decolorCells(transformationCells);
            if (addIn.theRibbon.dropDown1.Items.Count > 0) //if we have transformations available
            {
                FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == addIn.theRibbon.dropDown1.SelectedItem.Label);

                Excel.Range R = (addIn.Application.Selection);
                string formula = BBTransformations.RemoveFirstSymbol((string)R[1, 1].Formula);
                addIn.theRibbon.Preview.Text = T.ApplyOn(formula);
                addIn.theRibbon.valuePreview.Text = getValue((ExcelRaw.Range)R[1, 1].UnderlyingObject, addIn.theRibbon.Preview.Text);
                addIn.theRibbon.valuePreview.ShowImage = (addIn.theRibbon.valuePreview.Text != R[1, 1].Value.ToString());

                if (R.Count == 1)
                {
                    foreach (Excel.Worksheet worksheet in addIn.Application.Worksheets)
                    {
                        applyInRange(T, (ExcelRaw.Range)worksheet.UsedRange.UnderlyingObject, true);
                    }
                }
                else
                {
                    applyInRange(T, addIn.Application.Selection, true);
                }
            }
        }

        public void clearTransformationsRibbon(BBAddIn addIn)
        {
            addIn.theRibbon.Preview.Text = "";
            addIn.theRibbon.valuePreview.Text = "";
            addIn.theRibbon.valuePreview.ShowImage = false;
            addIn.theRibbon.dropDown1.Items.Clear();
            addIn.bbColorSmells.decolorCells(transformationCells);
        }
    }
}