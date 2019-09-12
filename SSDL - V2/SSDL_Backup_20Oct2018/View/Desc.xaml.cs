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
    /// Interaction logic for Desc.xaml
    /// </summary>
    public partial class Desc : Window
    {
        public bool cancel = false;
        public int maxAllowed = 30;
        public string remainingChars = " chars. remaining";
        public Desc()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            cancel = true;
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBoxdesc.MaxLength = maxAllowed;
            updateRemainingChars();
            btnOK.Focus();
        }

        private void updateRemainingChars()
        {
            lblRemainingText.Content = (maxAllowed - textBoxdesc.Text.Length).ToString() + remainingChars;
        }

        private void textBoxdesc_TextChanged(object sender, TextChangedEventArgs e)
        {
            updateRemainingChars();
        }
    }
}
