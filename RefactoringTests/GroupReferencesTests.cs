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
            //#if DEBUG
            //    excel.Visible = true;
            //#else
                excel.Visible = false;
            //#endif

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
        public void SingleCell()
        {
            test("SUM(A1)", "SUM(A1)");
        }

        [TestMethod]
        public void TwoCells()
        {
            test("SUM(A1,A2)","SUM(A1:A2)");
        }

        [TestMethod]
        public void ThreeCells()
        {
            test("SUM(A1,A2,A3)", "SUM(A1:A3)");
        }

        [TestMethod]
        public void Disconnected()
        {
            test("SUM(A1,A2,F4)", "SUM(A1:A2,F4)");
        }

        [TestMethod]
        public void TwoDisconnected()
        {
            test("SUM(A1,F5,F6,A2,F4)", "SUM(A1:A2,F4:F6)");
        }

        [TestMethod]
        public void Absolute()
        {
            test("SUM($A$1,$A$2,$A$3)", "SUM($A$1:$A$3)");
        }

        [TestMethod]
        public void MixedAbsolute()
        {
            test("SUM($A$1,A2,$A$3)", "SUM($A$1,$A$3,A2)");
        }

        [TestMethod]
        public void Test30()
        {
            // Application.Union has an argument limit of 30
            test("SUM(A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17,A18,A19,A20,A21,A22,A23,A24,A25,A26,A27,A28,A29,A30)",
                "SUM(A1:A30)");
        }

        [TestMethod]
        public void Test31()
        {
            // Application.Union has an argument limit of 30
            test("SUM(A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17,A18,A19,A20,A21,A22,A23,A24,A25,A26,A27,A28,A29,A30,A31)",
                "SUM(A1:A31)");
        }

        [TestMethod]
        public void Test32()
        {
            // Application.Union has an argument limit of 30
            test("SUM(A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17,A18,A19,A20,A21,A22,A23,A24,A25,A26,A27,A28,A29,A30,A31,A32)",
                "SUM(A1:A32)");
        }

        [TestMethod]
        public void TestCommaInRange()
        {
            var ranges = new [] {"A1", "A2"};
            var range = ws.Range[String.Join((string)excel.International[XlApplicationInternational.xlListSeparator], ranges)];
            Assert.AreEqual(String.Join(",", ranges), range.Address[false,false]);
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
