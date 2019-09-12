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

namespace pdf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public System.Collections.ObjectModel.ObservableCollection<Layerlist> _layerlisodwg;
        public System.Collections.ObjectModel.ObservableCollection<Layerlist> Layerlisodwg
        {
            get { return _layerlisodwg; }
            set
            {
                _layerlisodwg = value;                
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            this.dgvLayerlist.ItemsSource = Layerlisodwg;
            
        }
    }

    public class Layerlist
    {
        public bool include;
        public string Layername;
        public bool EntityAvail;
    }
}
