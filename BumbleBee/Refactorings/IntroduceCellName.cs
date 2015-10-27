using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using BumbleBee.Refactorings.Util;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using XLParser;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
//using XLParser;

namespace BumbleBee.Refactorings
{
    public class IntroduceCellName : RangeRefactoring
    {
        public override void Refactor(ExcelRaw.Range toName)
        {
            string subject = (toName.Count > 1) ? "Range" : "Cell";
            string name = Microsoft.VisualBasic.Interaction.InputBox($"{subject} name:", "Introduce Name", existingName(toName));
            Refactor(toName, name);
        }

        /// <summary>
        /// Name a cell, and replace all references to it with the name
        /// </summary>
        /// <exception cref="AggregateException">If any cells could not be inlined, with as innerexceptions the individual errors.</exception>
        public void Refactor(ExcelRaw.Range toName, string name)
        {
            var parse = Helper.Parse(name).SkipToRelevant();
            if (!parse.Is(GrammarNames.NamedRange))
            {
                throw new ArgumentException($"Name {name} is not a valid name, because Excel interpets it as a {parse.Type()}");
            }

            // Set the name
            var Worksheet = toName.Worksheet;
            var Names = Worksheet.Names;

            Names.Add(name, toName.Address[true, true]);

            Marshal.ReleaseComObject(Names);
            Marshal.ReleaseComObject(Worksheet);

            // Perform refactoring
            var ctx = toName.CreateContext();
            InlineFormula.RefactorSingle(toName, ctx, ctx.Parse(name));
        }

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.NonEmpty;

        private static string existingName(ExcelRaw.Range range)
        {
            try
            {
                ExcelRaw.Name eName = range.Name;
                string name = eName.Name;
                Marshal.ReleaseComObject(eName);
                return name;
            }
            catch (COMException)
            {
                return "";
            }
        }
    }
}