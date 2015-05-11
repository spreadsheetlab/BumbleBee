using System;
using System.Runtime.InteropServices;
using ExcelAddIn3.Refactorings;
using ExcelAddIn3.Refactorings.Util;
using Infotron.Parsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Excel;

namespace RefactoringTests
{
    [TestClass]
    public class GroupReferencesTests
    {
        private static Application excel;
        private static Workbook wb;
        private Worksheet ws;

        private IFormulaRefactoring testee;

        [ClassInitialize]
        public static void InitiateExcel(TestContext tc)
        {
            excel = new Application();
            excel.Visible = false;
            // excel.Visible = false;

            wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }

        [ClassCleanup]
        public static void TeardownExcel()
        {
            wb.Close(false);
            Marshal.FinalReleaseComObject(wb);
            wb = null;
            excel.Quit();
            Marshal.FinalReleaseComObject(excel);
            excel = null;
        }

        [TestInitialize]
        public void newSheet()
        {
            ws = (Worksheet)wb.Worksheets[1];
            testee = new GroupReferences(ws);
        }

        [TestCleanup]
        private void deleteSheet()
        {
            ws.Delete();
            Marshal.FinalReleaseComObject(ws);
            ws = null;
        }

        [TestMethod]
        public void TwoCells()
        {
            test("SUM(A1,A2)","SUM(A1:A2)");
        }

        private void test(string ungrouped, string grouped)
        {
            var ungroupedP = ungrouped.Parse();

            Assert.IsTrue(testee.CanRefactor(ungroupedP), "Should be able to refactor '{0}' but GroupReferences reported it cannot", ungrouped);

            var refactored = testee.Refactor(ungroupedP);
            Assert.AreEqual(Context.Empty.Parse(grouped), Context.Empty.ProvideContext(refactored));
        }
    }
}
