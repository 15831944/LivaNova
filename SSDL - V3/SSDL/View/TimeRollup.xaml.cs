using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SSDL.View
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Timerollup : Window
    {
        public int maxAllowed = 30;
        public string remainingChars = " chars. remaining";
        public System.Collections.ObjectModel.ObservableCollection<clsTimeRollUp> _timeRollup;
        public System.Collections.ObjectModel.ObservableCollection<clsTimeRollUp> timeRollupCol
        {
            get { return _timeRollup; }
            set
            {
                _timeRollup = value;
                //NotifyPropertyChanged();
            }
        }
        //= new ObservableCollection<clsTimeRollUp>();
        public bool readresult = false;
        public double totalTimeFromDwg = 0;
        public string drawingNumber = "";
        public double finalTimeFromUI = 0;
        public Timerollup()
        {
            InitializeComponent();
        }


        private void dgvTimeRollup_GotFocus(object sender, RoutedEventArgs e)
        {
            DataGridCell cell = e.OriginalSource as DataGridCell;
            if (cell != null && cell.Column is DataGridCheckBoxColumn)
            {
                dgvTimeRollup.BeginEdit();
                CheckBox chkBox = cell.Content as CheckBox;
                if (chkBox != null)
                {
                    chkBox.IsChecked = !chkBox.IsChecked;
                }
            }
            calcTotal();
        }

        void OnChecked(object sender, RoutedEventArgs e)
        {
            calcTotal();
        }

        private void calcTotal()
        {
            double _tempTotal = totalTimeFromDwg + timeRollupCol.Where(a => a.include == true).Sum(a => a.Time);
            _tempTotal = Math.Round(_tempTotal / 60, 2);//15-05-2018, added to change this to mins
            //TotalMins.Content = "Total : " + _tempTotal + " secs";//15-05-2018
            TotalMins.Content = "Total : " + _tempTotal + " mins";//15-05-2018
            finalTimeFromUI = _tempTotal;
        }

        private void updateRemainingChars()
        {
            lblRemainingText.Content = (maxAllowed - textBoxdesc.Text.Length).ToString() + remainingChars;
        }

        private void textBoxdesc_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateRemainingChars();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxdesc.MaxLength = maxAllowed;
            updateRemainingChars();
            dgvTimeRollup.ItemsSource = timeRollupCol;
            calcTotal();        
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Time Rollup for Drawing '" + drawingNumber + "' is '" + TotalMins.Content.ToString().Replace("Total: ", "").Trim() + "'. Do you wish to proceed?", "SSDL", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                readresult = true;
                this.Close();
            }
            else
            {
                readresult = false;
            }
        }

        private void ckIncludedbHeader_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ckIncludedbHeader_Checked(object sender, RoutedEventArgs e)
        {

            // timeRollupCol.Select(a => a.include = true);
            timeRollupCol.All(c => { c.include = true; return true; });
        }

        private void ckIncludedbHeader_Unchecked(object sender, RoutedEventArgs e)
        {
            //timeRollupCol.Select(a => a.include = false);
            timeRollupCol.All(c => { c.include = false; return true; });
        }
    }
    public class clsTimeRollUp : INotifyPropertyChanged
    {
        private bool _include { get; set; }
        private string _Category { get; set; }
        private int _Time { get; set; }

        public bool include { get { return _include; } set { _include = value; NotifyPropertyChanged(); } }
        public string Category { get { return _Category; } set { _Category = value; NotifyPropertyChanged(); } }
        public int Time { get { return _Time; } set { _Time = value; NotifyPropertyChanged(); } }


        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
