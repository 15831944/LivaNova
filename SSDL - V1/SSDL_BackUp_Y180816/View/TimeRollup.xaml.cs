using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Shapes;

namespace SSDL.View
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Timerollup : Window
    {
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
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
    }
    public class clsTimeRollUp
    {
        public bool include { get; set; }
        public string Category { get; set; }
        public int Time { get; set; }
    }
}
