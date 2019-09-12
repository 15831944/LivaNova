using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
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

namespace SSDL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DWGViewer.DWGViewer _viewer = new DWGViewer.DWGViewer();
        string path;
        List<string> Currentdrawinglist = new List<string>();
        private object dummyNode = null;
        public MainWindow()
        {
            InitializeComponent();
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            path = getpath.IniReadValue("FilePath", "DRAWING_PATH");
            //_viewer.Dock = System.Windows.Forms.DockStyle.Fill;
        
            host1.Child = _viewer;
            // _viewer.Dock = System.Windows.Forms.DockStyle.Fill;
            string preview = getpath.IniReadValue("FilePath", "PREVIEW_ON");
            if (preview.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase) || preview.Equals("YES", StringComparison.InvariantCultureIgnoreCase))
                chkPreview.IsChecked = true;
        }

        void folder_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == dummyNode)
            {
                item.Items.Clear();
                try
                {
                    foreach (string s in Directory.GetDirectories(item.Tag.ToString()))
                    {
                        TreeViewItem subitem = new TreeViewItem();
                        subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                        subitem.Tag = s;
                        subitem.FontWeight = FontWeights.Normal;
                        subitem.Items.Add(dummyNode);
                        subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                        item.Items.Add(subitem);
                    }
                }
                catch (Exception) { }
            }
        }


        private void BtnFind_Click(object sender, RoutedEventArgs e)
        {
            LstFiles.Items.Clear();
            string findtext = TxtFind.Text;

            List<string> getfindtext = Currentdrawinglist.AsEnumerable()
                                  .Where(x => x.ToString().Contains(findtext)).ToList();

            for (int i = 0; i < getfindtext.Count(); i++)
            {
                LstFiles.Items.Add(getfindtext[i]);
            }

            if (LstFiles.Items.Count == 0)
            {
                LstFiles.Items.Add("No Items Found");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LstFiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
            //if (LstFiles.SelectedIndex == -1) return;
            //string openpath = (string)item.Tag + "\\" + LstFiles.SelectedItem;
            //AcadFunctions.OpenDrawing(openpath);
            openDwg();
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            openDwg();
        }

        private void openDwg()
        {
            try
            {
                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                if (LstFiles.SelectedIndex == -1) return;
                string openpath = (string)item.Tag + "\\" + LstFiles.SelectedItem;
                AcadFunctions.OpenDrawing(openpath);
                BtnClose_Click(null, null);
            }
            catch { }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (string s in Directory.GetDirectories(path))
                {
                    TreeViewItem item = new TreeViewItem();
                    item.Header = s.Remove(0, path.Length);
                    item.Tag = s;
                    item.FontWeight = FontWeights.Normal;
                    item.Items.Add(dummyNode);
                    item.Expanded += new RoutedEventHandler(folder_Expanded);
                    foldersItem.Items.Add(item);
                }

                host1.Width = dpPreview.Width;
                host1.Height = dpPreview.Height;
                _viewer.Width = Convert.ToInt32(host1.Width);
                _viewer.Height = Convert.ToInt32(host1.Height);
            }
            catch { }
        }

        private void foldersItem_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            try
            {
                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                string[] filePaths = Directory.GetFiles((string)item.Tag, "*.dwg", SearchOption.TopDirectoryOnly);

                LstFiles.Items.Clear();
                Currentdrawinglist.Clear();

                for (int i = 0; i < filePaths.Count(); i++)
                {
                    LstFiles.Items.Add(System.IO.Path.GetFileName(filePaths[i]));
                    Currentdrawinglist.Add(System.IO.Path.GetFileName(filePaths[i]));
                }
            }
            catch { }
        }

        private void chkPreview_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));

                if (_viewer.IsDisposed)
                {
                    _viewer = new DWGViewer.DWGViewer();
                    host1.Child = _viewer;
                }

                if (chkPreview.IsChecked == true)
                {
                    getpath.IniWriteValue("FilePath", "PREVIEW_ON", "TRUE");
                    _viewer.Visible = true;
                    LstFiles_SelectionChanged(null, null);
                }
                else
                {
                    getpath.IniWriteValue("FilePath", "PREVIEW_ON", "FALSE");
                    _viewer.Visible = false;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void BtnRef_Click(object sender, RoutedEventArgs e)
        {

            foreach (TreeViewItem item in foldersItem.Items)
                CollapseTreeviewItems(item);
        }

        public void CollapseTreeviewItems(TreeViewItem Item)
        {
            Item.IsExpanded = false;

            foreach (TreeViewItem item in Item.Items)
            {
                try
                {
                    item.IsExpanded = false;

                    if (item.HasItems)
                        CollapseTreeviewItems(item);
                }
                catch { }

            }
        }

        private void LstFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                if (LstFiles.SelectedItem != null && item != null)
                {
                    setPreviewImages(System.IO.Path.Combine((string)item.Tag, LstFiles.SelectedItem.ToString()));

                }

            }
            catch { }
        }

        private void setPreviewImages(string fullFileName, bool emptyImage = false)
        {
            try
            {
                if (!emptyImage)
                {
                    if (chkPreview.IsChecked == true)// _bitMap != null)
                    {
                        if (_viewer.IsDisposed)
                        {
                            _viewer = new DWGViewer.DWGViewer();
                            host1.Child = _viewer;
                        }
                        _viewer.loadFile(fullFileName);
                        return;
                    }
                    else
                    {
                        if (!_viewer.IsDisposed)
                            _viewer.Dispose();
                    }
                }
                else
                {
                    if (!_viewer.IsDisposed)
                        _viewer.Dispose();
                }
            }
            catch (Exception ex) { }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _viewer.Dispose();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            //if (_viewer.IsDisposed)
            //{
            //    _viewer = new DWGViewer.DWGViewer();
            //    host1.Child = _viewer;
            //}
           
        }
    }
}
