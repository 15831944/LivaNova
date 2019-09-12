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
using System.Windows.Shapes;

namespace SSDL.View
{
    /// <summary>
    /// Interaction logic for ProgressBar.xaml
    /// </summary>
    public partial class ProgressBar : Window
    {
        public ProgressBar(string labelName= "Publish in progress")
        {
            InitializeComponent();
            labelContent = labelName;
        }
        public string labelContent { get; set; }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
           // lblContent.Content = labelContent;
        }
    }
}
