﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
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
using System.Windows.Forms;
using BumbleBee.Refactorings;
using BumbleBee.TaskPanes;
using FSharpEngine;
using Irony.Parsing;
using XLParser;

namespace BumbleBee
{
    public enum ApplyTo
    {
        Range,
        Worksheet,
        Workbook
    }

    public partial class BBAddIn
    {
        private Excel.Application netofficeApplication;

        public Excel.Application App => (netofficeApplication = netofficeApplication ?? new Excel.Application(null, Globals.BBAddIn.Application));

        public BumbleBeeRibbon theRibbon;

        public BBColorSmells bbColorSmells { get; private set; }
        public BBTransformations bbTransformations { get; private set; }
        public BBMenuRefactorings bbMenuRefactorings { get; private set; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            theRibbon = new BumbleBeeRibbon();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { theRibbon });
            
        }

        public void InitializeBB()
        {
            bbTransformations.startsTransformationRules();
        }

        public void Log(string LogMessage)
        {
            string currentWorkbookFilePath = this.Application.ActiveWorkbook.Path;
            string LogFileName = "spreadsheets.log";
            string LogFile = System.IO.Path.Combine(currentWorkbookFilePath, LogFileName);
            var file = new System.IO.StreamWriter(LogFile, true);
            file.WriteLine(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ", " + LogMessage);
            file.Close();
        }

        public ExcelRaw.Worksheet GetWorksheetByName(string name)
        {
            return Application.Worksheets
                .Cast<ExcelRaw.Worksheet>()
                .FirstOrDefault(worksheet => worksheet.Name == name);
        }

        // TODO: Better place / dynamic location, preferably inside source control
        private readonly string[] BumbleBeeDebugStartupfiles =
        {
            @"C:\bumblebee\startup.xlsx",
            @"C:\bumblebee\startup.xlsm"
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            bbColorSmells = new BBColorSmells(this);
            bbTransformations = new BBTransformations(this);
            bbMenuRefactorings = new BBMenuRefactorings(this);

            #if DEBUG
            foreach (var startupfile in BumbleBeeDebugStartupfiles.Where(System.IO.File.Exists))
            {
                Application.Workbooks.Open(startupfile);
            }
            #endif

            bbMenuRefactorings.startBBRefactorings();
        }


        void Application_WorkbookOpen(ExcelRaw.Workbook Wb)
        {
            InitializeBB();
        }

        private readonly BBMenuRefactorings menuRefactorings;

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            Application.WorkbookOpen += new ExcelRaw.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            Application.SheetSelectionChange += new ExcelRaw.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Application.WorkbookAfterSave += new ExcelRaw.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
        }
        #endregion

        void Application_SheetSelectionChange(object Sh, ExcelRaw.Range Target)
        {
            bbTransformations.clearTransformationsRibbon(this);
        }

        void Application_WorkbookAfterSave(Microsoft.Office.Interop.Excel.Workbook w, bool success)
        {
            InitializeBB();
        }
    }
}
