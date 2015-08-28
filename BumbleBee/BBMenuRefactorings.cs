using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using BumbleBee.Refactorings;
using BumbleBee.TaskPanes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

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
            new BBMenuRefactorings.RefactorMenuItem {MenuText="+ to SUM",Refactoring=new OpToAggregate()},
            new BBMenuRefactorings.RefactorMenuItem {MenuText="SUM to SUMIF", Refactoring=new GroupReferences()},
            new BBMenuRefactorings.RefactorMenuItem {MenuText="Group References", Refactoring=new GroupReferences(), NewGroup = true},
            new BBMenuRefactorings.RefactorMenuItem {MenuText="Inline Formula", Refactoring=new InlineFormula(), NewGroup = true},
            new BBMenuRefactorings.RefactorMenuItem {MenuText="Extract Formula", Refactoring=new ExtractFormulaMenuStub()},
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
                            String.Format(
                                "Errors:\n{0}",
                                String.Join("\n\n", e.InnerExceptions.Select(ie => ie.Message))
                                )
                            );
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(string.Format("Error: {0}", e.Message));
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
    }
}