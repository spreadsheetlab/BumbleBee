using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Infotron.FSharpFormulaTransformation;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace ExcelAddIn3.UserDialogs
{
    /// <summary>
    /// Interaction logic for ExtractFormulaDialog.xaml
    /// </summary>
    public partial class ExtractFormulaDialog : Window
    {
        private readonly string orFormula;

        private string _formula;
        public string Formula
        {
            get { return _formula; }
            set
            {
                if (!orFormula.Contains(value))
                {
                    throw new ArgumentException("Original formula does not contain subformula.");
                }
                /*
                FSharpTransformationRule T = new FSharpTransformationRule();
                if (T.ParseToTree(value) == null)
                {
                    throw new ArgumentException("Not a valid formula.");
                }
                 */
                _formula = value;
            }
        }

        public BBAddIn.ExtractDirection Direction { get; set; }

        readonly static Regex CellAddressRegex = new Regex("[A-Z]+\\d+");
        private string _cellAddress;
        public string CellAddress
        {
            get { return _cellAddress; }
            set
            {
                if (!CellAddressRegex.IsMatch(value))
                {
                    throw new ArgumentException("Invalid cell address");
                }
                _cellAddress = value;
            }
        }

        public ExtractFormulaDialog(Range from)
        {
            InitializeComponent();
            this.DataContext = this;

            orFormula = from.Formula.Substring(1);
            Formula = orFormula;
            CellAddress = from.Address[false, false];
            
            // Make sure only valid directions can be chosen and a sensible default is provided
            radioDirLeft.IsEnabled = from.Column > 1;
            radioDirUp.IsEnabled = from.Row > 1;
            radioDirUp.IsChecked = !radioDirLeft.IsEnabled && radioDirUp.IsEnabled;
            radioDirRight.IsChecked = !radioDirLeft.IsEnabled && !radioDirUp.IsEnabled;
        }

        private void buttonDialogExtractFormula_Click(object sender, RoutedEventArgs e)
        {
            checkDirection();
            this.DialogResult = true;
        }

        private void checkDirection()
        {
            if (radioDirUp.IsChecked == true) {
                Direction = BBAddIn.ExtractDirection.Up;
            } else if (radioDirDown.IsChecked == true) {
                Direction = BBAddIn.ExtractDirection.Down;
            } else if (radioDirLeft.IsChecked == true) {
                Direction = BBAddIn.ExtractDirection.Left;
            } else if (radioDirRight.IsChecked == true) {
                Direction = BBAddIn.ExtractDirection.Right;
            } else if (RadioDirFixed.IsChecked == true) {
                Direction = BBAddIn.ExtractDirection.Fixed;
            }
        }
    }
}
