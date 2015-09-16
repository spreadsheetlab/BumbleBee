using System;
using System.Collections.Generic;
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
using System.ComponentModel;
using Infotron.Util;
using Infotron.Parsing;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using BumbleBee.Refactorings;
using BumbleBee.Refactorings.Util;

namespace BumbleBee.TaskPanes
{
    /// <summary>
    /// Interaction logic for ExtractFormulaTaskPane.xaml
    /// </summary>
    public partial class ExtractFormulaTaskPane : UserControl, INotifyPropertyChanged
    {

        public ExtractFormulaTaskPane()
        {
            InitializeComponent();
            DataContext = this;

            // Update the refactored formula when address or subformula change
            PropertyChanged += (e, o) =>
            {
                checkValid();
                switch (o.PropertyName)
                {
                    case "orFormula":
                    case "Formula":
                    case "NewCellAddress":
                        setRefactoredFormula();
                        break;
                }
            };
        }

        
        private ContextNode orFormula;

        public string orFormulaStr
        {
            get { return orFormula != null ? orFormula.Print() : ""; }
        }

        public void newRange(Range from)
        {
            OrRange = from;
            var topleft = from.TopLeft();

            try
            {
                orFormula = Helper.ParseCtx(topleft);
            }
            catch (InvalidDataException)
            {
                orFormula = topleft.CreateContext().Parse("0");
            }
            finally
            {
                OnPropertyChanged("orFormula");
                OnPropertyChanged("orFormulaStr");
            }

            radioDirUp.IsEnabled = true;
            radioDirDown.IsEnabled = true;
            radioDirLeft.IsEnabled = true;
            radioDirRight.IsEnabled = true;

            // Disable up-down unless all selected areas are single-row
            var singlerow = from.Areas.Cast<Range>().All(area => {
                 var row = area.TopLeft().Row;
                 return area.Cells.Cast<Range>().All(c => c.Row == row);
            });
            radioDirUp.IsEnabled = radioDirUp.IsEnabled && singlerow;
            radioDirDown.IsEnabled = radioDirDown.IsEnabled && singlerow;

            // Disable left-right unless all selected areas are single-column
            var singlecolumn = from.Areas.Cast<Range>().All(area =>
            {
                var col = area.TopLeft().Column;
                return area.Cells.Cast<Range>().All(c => c.Column == col);
            });
            radioDirLeft.IsEnabled = radioDirLeft.IsEnabled && singlecolumn;
            radioDirRight.IsEnabled = radioDirRight.IsEnabled && singlecolumn;
                
            // Provide a default direction checkbox
            radioDirLeft.IsChecked = radioDirLeft.IsEnabled && RadioDirFixed.IsChecked == false;
            radioDirUp.IsChecked = radioDirUp.IsEnabled && RadioDirFixed.IsChecked != true
                                   && radioDirLeft.IsChecked != true;
            radioDirRight.IsChecked = radioDirRight.IsEnabled && RadioDirFixed.IsChecked != true
                                   && radioDirLeft.IsChecked != true
                                   && radioDirUp.IsChecked != true;
            radioDirDown.IsChecked = radioDirDown.IsEnabled && RadioDirFixed.IsChecked != true
                                   && radioDirLeft.IsChecked != true
                                   && radioDirUp.IsChecked != true
                                   && radioDirRight.IsChecked != true;
            RadioDirFixed.IsChecked = RadioDirFixed.IsChecked == true || (
                                      RadioDirFixed.IsEnabled
                                   && radioDirLeft.IsChecked != true
                                   && radioDirUp.IsChecked != true
                                   && radioDirRight.IsChecked != true
                                   && radioDirDown.IsChecked != true);
            
            checkDirection();
        }

        public void init(Range from)
        {
            newRange(from);
            FormulaStr = orFormula.Print();
        }

        private Range orRange;

        private Range OrRange
        {
            get { return orRange; }
            set
            {
                orRange = value;
                OnPropertyChanged("OrCellAddressStr");
            }
        }
        private Range TopLeft { get { return OrRange != null ? OrRange.TopLeft() : null; } }


        public string OrCellAddressStr
        {
            get { return OrRange != null ? TopLeft.Address[false,false] : null; }
        }

        private Location newCellAddress;

        public Location NewCellAddress
        {
            get
            {
                return newCellAddress;
            }
            private set
            {
                newCellAddress = value;
                OnPropertyChanged("NewCellAddress");
                OnPropertyChanged("NewCellAddressStr");
                OnPropertyChanged("FixedAddressStr");
            }
        }

        public string NewCellAddressStr
        {
            get { return NewCellAddress != null ? NewCellAddress.ToString() : ""; }
        }

        private bool fixedAddressValid = true;
        public string FixedAddressStr
        {
            get { return NewCellAddress != null ? NewCellAddress.Address() : ""; }
            set
            {
                try
                {

                    if (!Helper.isValidAddress(value))
                    {
                        throw new ArgumentException("Invalid cell address");
                    }
                    NewCellAddress = new Location(value) {RowFixed = true, ColumnFixed = true};
                    fixedAddressValid = true;
                }
                catch (Exception)
                {
                    fixedAddressValid = false;
                    throw;
                }
                finally
                {
                    checkValid();
                }
            }
        }


        public ContextNode RefactoredFormula { get; private set; }
        public string RefactoredFormulaStr => RefactoredFormula != null ? RefactoredFormula.Print() : "";

        private void setRefactoredFormula()
        {
            if (orFormula == null || Formula == null || NewCellAddress == null) return;
            RefactoredFormula = orFormula.Replace(Formula, orFormula.Ctx.Parse(NewCellAddress.Address()));
            OnPropertyChanged("RefactoredFormula");
            OnPropertyChanged("RefactoredFormulaStr");
        }

        public ContextNode Formula { get; private set; }

        private bool formulaValid = true;
        public string FormulaStr
        {
            get { return Formula != null ? Formula.Print() : ""; }
            set
            {
                try
                {
                    try
                    {
                        var f = orFormula.Ctx.Parse(value);
                        if (!orFormula.Contains(f))
                        {
                            throw new ArgumentException("Original formula does not contain subformula.", nameof(value));
                        }
                        Formula = f;
                    }
                    catch (InvalidDataException e)
                    {
                        throw new ArgumentException("Not a valid formula", nameof(value), e);
                    }
                    formulaValid = true;
                    OnPropertyChanged("Formula");
                    OnPropertyChanged("FormulaStr");
                }
                catch (Exception)
                {
                    formulaValid = false;
                    throw;
                }
                finally
                {
                    checkValid();
                }
            }
        }

        public ExtractFormula.Direction Direction { get; set; }

        private void checkDirection(object sender = null, RoutedEventArgs e = null)
        {
            if (RadioDirFixed == null || RadioDirFixed.IsChecked == true)
            {
                Direction = null;
            }
            else if (radioDirUp.IsChecked == true)
            {
                Direction = ExtractFormula.Direction.Up;
            }
            else if (radioDirDown.IsChecked == true)
            {
                Direction = ExtractFormula.Direction.Down;
            }
            else if (radioDirLeft.IsChecked == true)
            {
                Direction = ExtractFormula.Direction.Left;
            }
            else if (radioDirRight.IsChecked == true)
            {
                Direction = ExtractFormula.Direction.Right;
            }
            if (Direction != null && OrRange != null)
            {
                try
                {
                    NewCellAddress = new Location(
                        OrRange
                            .Offset[Direction.RowOffset, Direction.ColOffset]
                            .TopLeft()
                            .Address[false, false]
                        );
                }
                catch (COMException)
                {
                    // column or row 0 and offset was left or up
                    // Not really a good solution on what to preview, just show the original
                    NewCellAddress = new Location(OrRange.TopLeft().Address[false,false]);
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        private void buttonDialogExtractFormula_Click(object sender, RoutedEventArgs e)
        {
            doExtract();
        }

        public void doExtract()
        {
            try
            {
                // Fixed location
                if (Direction == null)
                {
                    ExtractFormula.Refactor(OrRange, NewCellAddress, Formula);
                }
                else
                {
                    ExtractFormula.Refactor(OrRange, Direction, Formula);
                }
                Globals.BBAddIn.bbMenuRefactorings.extractFormulaCtp.Visible = false;
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkValid()
        {
            buttonExtractFormula.IsEnabled = formulaValid && fixedAddressValid;
        }
    }        
}
