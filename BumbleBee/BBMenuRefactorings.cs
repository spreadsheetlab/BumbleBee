using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BumbleBee.Refactorings;
using BumbleBee.Refactorings.Util;
using BumbleBee.TaskPanes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;

namespace BumbleBee
{
    public class BBMenuRefactorings
    {
        /// Enable to profile speed where implemented
        internal const bool PROFILE = false;

        private BBAddIn addIn;
        internal TaksPaneWPFContainer<ExtractFormulaTaskPane> extractFormulaTp;
        internal CustomTaskPane extractFormulaCtp;

        private readonly List<RefactorMenuItem> contextMenuRefactorings = new List<BBMenuRefactorings.RefactorMenuItem>
        {
            new RefactorMenuItem {MenuText="Change to SUM or SUMIF",Refactoring=new PlusToSumMenuStub()},
            new RefactorMenuItem {MenuText="Group References", Refactoring=new GroupReferences(), NewGroup = true},
            new RefactorMenuItem {MenuText="Extract Formula", Refactoring=new ExtractFormulaMenuStub(), NewGroup = true},
            new RefactorMenuItem {MenuText="Inline Formula", Refactoring=new InlineFormula()},
            new RefactorMenuItem {MenuText="Introduce Cell Name", Refactoring = new IntroduceCellName(), NewGroup = true },
            #if DEBUG
            new RefactorMenuItem {MenuText="[DEBUG] Op to Aggregate", Refactoring = new OpToAggregate(), NewGroup = true},
            new RefactorMenuItem {MenuText="[DEBUG] Introduce Conditional Aggregate", Refactoring=new AgregrateToConditionalAggregrate()},
            #endif
        };

        public BBMenuRefactorings(BBAddIn addIn)
        {
            this.addIn = addIn;
        }

        public void startBBRefactorings()
        {
            extractFormulaTp = new TaksPaneWPFContainer<ExtractFormulaTaskPane>(new ExtractFormulaTaskPane());
            extractFormulaCtp = addIn.CustomTaskPanes.Add(extractFormulaTp, "Extract formula");

            RefactoringContextMenuInitialize();
            addIn.Application.SheetBeforeRightClick +=
                new AppEvents_SheetBeforeRightClickEventHandler(RefactorMenuEnableRelevant);
        }

        private class RefactorMenuItem
        {
            public string MenuText { get; set; }
            public IRangeRefactoring Refactoring { get; set; }
            public bool NewGroup { get; set; }
            public Microsoft.Office.Core.CommandBarButton Button { get; set; }
        }

        
        /// <summary>
        /// Create the refactoring context menu
        /// </summary>
        private void RefactoringContextMenuInitialize()
        {
            const string tag = "REFACTORMENU";
            var cellcontextmenu = addIn.Application.CommandBars["Cell"];

            // Check if menu already defined
            if (cellcontextmenu.FindControl(Microsoft.Office.Core.MsoControlType.msoControlPopup, 0, tag) != null) return;

            var menu = (Microsoft.Office.Core.CommandBarPopup)cellcontextmenu.Controls.Add(
                Type:Microsoft.Office.Core.MsoControlType.msoControlPopup,
                Before: cellcontextmenu.Controls.Count,
                Temporary: true);
            menu.Caption = "Refactor";
            menu.BeginGroup = true;
            menu.Tag = tag;

            foreach (var menuitem in contextMenuRefactorings)
            {
                var control = (Microsoft.Office.Core.CommandBarButton) menu.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Temporary:true);
                control.Caption = menuitem.MenuText;
                control.BeginGroup = menuitem.NewGroup;
                // Disable by default, only enable when relevant
                control.Enabled = false;
                menuitem.Button = control;
                // Create a copy of the iterator item because we're going to use it within a closure
                var refactoring = menuitem.Refactoring;
                control.Click += (Microsoft.Office.Core.CommandBarButton ctrl, ref bool cancelDefault) =>
                {
                    try
                    {
                        refactoring.Refactor(addIn.Application.Selection);
                    }
                    catch (AggregateException e)
                    {
                        MessageBox.Show(
                            $"Errors:\n{String.Join("\n\n", e.InnerExceptions.Select(ie => ie.Message))}"
                            );
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show($"Error: {e.Message}");
                    }
                };
            }
        }


        private const int MAX_NUMBER_OF_CELLS_TO_CHECK = 50;

        /// <summary>
        /// This method enables/disables the refactorings in the context menu
        /// </summary>
        private void RefactorMenuEnableRelevant(object sheet, ExcelRaw.Range target, ref bool cancel)
        {
            // Because of performance reasons, don't check if a large number of cells is selected
            var count = target.Count;
            foreach (var item in contextMenuRefactorings)
            {
                item.Button.Enabled = item.Refactoring.CanRefactor(target);
            }
            Marshal.ReleaseComObject(target);
            if (Marshal.IsComObject(sheet)) Marshal.ReleaseComObject(sheet);
        }
    }
}