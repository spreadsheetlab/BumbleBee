#define DEBUG

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace BumbleBee
{
    public partial class BumbleBeeRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.FindApplicableTransformations();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.ApplyTransformation(ApplyTo.Range);
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.MakePreview(Globals.BBAddIn);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.ApplyTransformation(ApplyTo.Workbook);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.ApplyTransformation(ApplyTo.Worksheet);
        }

        #if DEBUG 

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbColorSmells.ColorSmells();
        }

        private void selectSmellType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbColorSmells.SelectSmellsOfType();
        }

        private void buttonInitializeBumbleBee_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.bbTransformations.AddSheetBumbleBeeTransformations();
        }

        #endif
    }
}
