using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
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
using System.ComponentModel;

namespace ExcelAddIn3.UserDialogs
{
    /// <summary>
    /// Interaction logic for ExtractFormulaDialog.xaml
    /// </summary>
    public partial class ExtractFormulaDialog : Window, INotifyPropertyChanged
    {
        private readonly Range orCell;
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
                if (!RefactoringHelper.isValidFormula(value))
                {
                    throw new ArgumentException("Not a valid formula");
                }
                _formula = value;
                OnPropertyChanged("Formula");
            }
        }

        public RefactoringHelper.Direction Direction { get; set; }

        private string _cellAddress;
        public string CellAddress
        {
            get { return _cellAddress; }
            set
            {
                if (!RefactoringHelper.isValidAddress(value))
                {
                    throw new ArgumentException("Invalid cell address");
                }
                _cellAddress = value;
                OnPropertyChanged("CellAddress");
            }
        }

        public string RefactoredFormula { get; private set; }

        private void setRefactoredFormula()
        {
            // TODO: Move cell address if left/right/down/up is selected
            RefactoredFormula = RefactoringHelper.replaceSubFormula(orFormula, Formula, orCell.Address[true, true]);
            OnPropertyChanged("RefactoredFormula");
        }

        public ExtractFormulaDialog(Range from)
        {
            InitializeComponent();
            this.DataContext = this;

            orCell = from;
            orFormula = from.Formula.Substring(1);
            Formula = orFormula;
            CellAddress = from.Address[false, false];

            // Make sure only valid directions can be chosen and a sensible default is provided
            radioDirLeft.IsEnabled = from.Column > 1;
            radioDirUp.IsEnabled = from.Row > 1;
            radioDirUp.IsChecked = !radioDirLeft.IsEnabled && radioDirUp.IsEnabled;
            radioDirRight.IsChecked = !radioDirLeft.IsEnabled && !radioDirUp.IsEnabled;

            // Update the refactored formula when address or subformula change
            PropertyChanged += (e, o) => {
                if (o.PropertyName == "Formula" || o.PropertyName == "CellAddress")
                    setRefactoredFormula();
            };
            setRefactoredFormula();
        }

        private void buttonDialogExtractFormula_Click(object sender, RoutedEventArgs e)
        {
            checkDirection();
            this.DialogResult = true;
        }

        private void checkDirection()
        {
            if (radioDirUp.IsChecked == true) {
                Direction = RefactoringHelper.Direction.Up;
            } else if (radioDirDown.IsChecked == true) {
                Direction = RefactoringHelper.Direction.Down;
            } else if (radioDirLeft.IsChecked == true) {
                Direction = RefactoringHelper.Direction.Left;
            } else if (radioDirRight.IsChecked == true) {
                Direction = RefactoringHelper.Direction.Right;
            } else if (RadioDirFixed.IsChecked == true) {
                Direction = RefactoringHelper.Direction.Fixed;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
