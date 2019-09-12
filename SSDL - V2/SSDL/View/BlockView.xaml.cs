using System;
using System.Collections.Generic;
using System.IO;
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
using System.Configuration;
using System.Reflection;
using DWGViewer;

namespace SSDL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class BlockView : Window
    {
        DWGViewer.DWGViewer _viewer = new DWGViewer.DWGViewer();
        string path;
        List<string> Currentdrawinglist = new List<string>();
        private object dummyNode = null;
        public BlockView()
        {
            InitializeComponent();
            try
            {
                host1.Child = _viewer;
                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
                path = getpath.IniReadValue("FilePath", "SYMBOL_LIBRARY");
                string preview = getpath.IniReadValue("FilePath", "PREVIEW_ON");
                if (preview.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase) || preview.Equals("YES", StringComparison.InvariantCultureIgnoreCase))
                    chkPreview.IsChecked = true;

                string windowclose = getpath.IniReadValue("FilePath", "CLOSE_SYMBOL_DIALOG_AFTER_INSERTION");
                if (preview.Equals("YES", StringComparison.InvariantCultureIgnoreCase))
                    chkclosewindow.IsChecked = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }

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


        private void setPreviewImages(string fullFileName, bool emptyImage = false)
        {
            try
            {
                if (!emptyImage)
                {
                    //System.Drawing.Bitmap _bitMap = ThumbnailReader.GetBitmap(fullFileName);
                    if (chkPreview.IsChecked == true)// _bitMap != null)
                    {
                        //ImageSourceConverter c = new ImageSourceConverter();

                        if (_viewer.IsDisposed)
                        {
                            _viewer = new DWGViewer.DWGViewer();
                            host1.Child = _viewer;
                        }
                        //_viewer.Visible = true;
                        _viewer.loadFile(fullFileName);
                        //image.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(_bitMap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                        //image.Source = (ImageSource)c.ConvertFrom(_bitMap);
                        return;
                    }
                    else
                    {
                        if (!_viewer.IsDisposed)
                            _viewer.Dispose();
                        // _viewer.Visible = false;
                    }
                }
                else
                {
                    if (!_viewer.IsDisposed)
                        _viewer.Dispose();
                    //_viewer.Visible = false;
                    // image.Source = new BitmapImage(new Uri("/SSDL;component/Empty.jpg", UriKind.Relative));
                }
            }
            catch (Exception ex) { }
            //image.Source = default image

            if (LstFiles.Items.Count > 0)
            {
                if (LstFiles.SelectedItems.Count <= 0)
                    BtnOpen.IsEnabled = true;
                else
                    BtnOpen.IsEnabled = true;
            }
            else
                BtnOpen.IsEnabled = true;
        }


        private void BtnFind_Click(object sender, RoutedEventArgs e)
        {
            try
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
                Imageview();
            }
            catch { }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void LstFiles_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                if (LstFiles.SelectedIndex == -1) return;
                string openpath = (string)item.Tag + "\\" + LstFiles.SelectedItem;
                AcadFunctions.OpenDrawing(openpath);
                this.Close();
            }
            catch { }
        }

        //Insert
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(Txtquantity.Value.ToString()))
                {
                    MessageBox.Show("Enter Quantity");
                    return;
                }
                //else if (string.IsNullOrEmpty(TxtItemno.Text))
                //{
                //    MessageBox.Show("Enter Item No");
                //    return;
                //}
                try
                {
                    int qty = Convert.ToInt32(Txtquantity.Value.ToString());
                    if (qty <= 0)
                    {
                        MessageBox.Show("Enter Valid Quantity");
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Enter Valid Quantity");
                    return;
                }

                //try
                //{
                //    int qty = Convert.ToInt32(TxtItemno.Text);
                //    if (qty <= 0)
                //    {
                //        MessageBox.Show("Enter Valid Item No");
                //        return;
                //    }
                //}
                //catch
                //{
                //    MessageBox.Show("Enter Valid Item No");
                //    return;
                //}

                if (TxtLen.Visibility == Visibility.Visible)
                {
                    try
                    {
                        double qty = Convert.ToDouble(TxtLen.Text);
                        if (qty <= 0)
                        {
                            MessageBox.Show("Enter Valid Length");
                            return;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Enter Valid Length");
                        return;
                    }
                }



                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                if (item != null)
                {
                    if (LstFiles.SelectedIndex == -1) return;
                    if (LstFiles.SelectedItem.ToString().StartsWith("No Record", StringComparison.InvariantCultureIgnoreCase))
                        return;

                    string openpath = (string)item.Tag + "\\" + LstFiles.SelectedItem;
                    AcadFunctions.InsertBlk(openpath, TxtPartno.Text, Txtquantity.Value.ToString(), TxtItemno.Text, TxtLen.Text);
                    //this.Close();
                }
            }
            catch { }

            // CLOSE_SYMBOL_DIALOG_AFTER_INSERTION = yes
            //IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            //string checkresult = getpath.IniReadValue("FilePath", "CLOSE_SYMBOL_DIALOG_AFTER_INSERTION");

            // if (checkresult.Equals("YES", StringComparison.InvariantCultureIgnoreCase))
            if (chkclosewindow.IsChecked == true)
            {
                this.Close();
            }

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
            }
            catch (Exception ex) { MessageBox.Show("Folder not loaded -" + ex.ToString()); }

            try
            {
                loadDropDown();
            }
            catch { MessageBox.Show("MDB error"); }

            try
            {
                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
                string checkresult = getpath.IniReadValue("FilePath", "CLOSE_SYMBOL_DIALOG_AFTER_INSERTION");
                if (checkresult.Equals("YES", StringComparison.InvariantCultureIgnoreCase))
                {
                    chkclosewindow.IsChecked = true;
                }
                else
                    chkclosewindow.IsChecked = false;
            }
            catch { }
        }

        private void loadDropDown()
        {
            try
            {
                //DatabaseClass _cls = new DatabaseClass();
                System.Data.DataTable dtProductsRM = DatabaseClass.GetProductsRM();
                if (dtProductsRM.Columns.Count > 0 && dtProductsRM.Rows.Count > 0)
                {
                    //CmbRMDesc.ItemsSource = dtProductsRM.DefaultView;
                    //CmbRMDesc.DisplayMemberPath = "RM_Desc";
                    //CmbRMDesc.SelectedValuePath = "RM_No";

                    //CmbVendor.ItemsSource = dtProductsRM.DefaultView;
                    //CmbVendor.DisplayMemberPath = "Vendor_Num";
                    //CmbVendor.SelectedValuePath = "RM_No";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void foldersItem_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
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

            Imageview();
        }

        private void LstFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string partno;
            string qty, itemno, len;
            try
            {
                TreeViewItem item = foldersItem.SelectedItem as TreeViewItem;
                if (LstFiles.SelectedItem != null && item != null)
                {
                    bool lenthcheck = false;
                    setPreviewImages(System.IO.Path.Combine((string)item.Tag, LstFiles.SelectedItem.ToString()));
                    AcadFunctions.ExtractObjectsFromFile(System.IO.Path.Combine((string)item.Tag, LstFiles.SelectedItem.ToString()), out partno, out qty, out itemno, out len, out lenthcheck);

                    TxtPartno.Text = partno;
                    //TxtItemno.Text = itemno.ToString();

                    Txtquantity.Value = 1;
                    // if (string.IsNullOrEmpty(qty))
                    // Txtquantity.Text = "";
                    // else
                    //  Txtquantity.Text = qty.ToString();

                    if (lenthcheck)
                    {
                        lbllen.Visibility = Visibility.Visible;
                        TxtLen.Visibility = Visibility.Visible;
                        TxtLen.Text = len;
                    }
                    else
                    {
                        lbllen.Visibility = Visibility.Hidden;
                        TxtLen.Visibility = Visibility.Hidden;
                        TxtLen.Text = "";
                    }
                }
                else
                {
                    Imageview();
                }
            }
            catch { }
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
                item.IsExpanded = false;

                try
                {
                    item.IsExpanded = false;

                    if (item.HasItems)
                        CollapseTreeviewItems(item);
                }
                catch { }
            }
            Imageview();
        }

        public void Imageview()
        {
            try
            {
                if (LstFiles.Items == null)
                    setPreviewImages("", true);
                //setPreviewImages("/SSDL;component/Empty.jpg");
                else if (LstFiles.Items.Count <= 0)
                    setPreviewImages("", true);
                //setPreviewImages(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Empty.jpg"));
                //setPreviewImages("/SSDL;component/Empty.jpg");
                else if (LstFiles.SelectedItems.Count <= 0)
                    setPreviewImages("", true);
                //  setPreviewImages(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Empty.jpg"));
                //  setPreviewImages("/SSDL;component/Empty.jpg");
            }
            catch { }
        }

        private void CmbRMDesc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (CmbRMDesc.SelectedValue != null && CmbRMDesc.SelectedIndex >= -1)
            //    findBasedOnCMB(CmbRMDesc.SelectedValue.ToString());
        }

        /// <summary>
        /// send the dropdown selected value(which is the file name)
        /// </summary>
        /// <param name="fileNameWithoutExtension">Padding not required</param>
        private void findBasedOnCMB(string fileNameWithoutExtension)
        {
            fileNameWithoutExtension = fileNameWithoutExtension.PadLeft(9, '0');
            List<string> getfindtext = Currentdrawinglist.AsEnumerable()
                                 .Where(x => x.ToString().StartsWith(fileNameWithoutExtension, StringComparison.InvariantCultureIgnoreCase)).ToList();
            LstFiles.Items.Clear();
            for (int i = 0; i < getfindtext.Count(); i++)
            {
                LstFiles.Items.Add(getfindtext[i]);
            }

            if (LstFiles.Items.Count == 0)
            {
                LstFiles.Items.Add("No Items Found");
            }
            Imageview();
        }

        private void CmbVendor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //if (CmbVendor.SelectedValue != null && CmbVendor.SelectedIndex >= -1)
            //    findBasedOnCMB(CmbVendor.SelectedValue.ToString());
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

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _viewer.Dispose();
        }

        private void chkclosewindow_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));

                if (chkclosewindow.IsChecked == true)
                {
                    getpath.IniWriteValue("FilePath", "CLOSE_SYMBOL_DIALOG_AFTER_INSERTION", "yes");
                }
                else
                {
                    getpath.IniWriteValue("FilePath", "CLOSE_SYMBOL_DIALOG_AFTER_INSERTION", "no");
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void Txtquantity_ValueChanged(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Txtquantity.Value <= 0)
                    Txtquantity.Value = 1;
            }
            catch { Txtquantity.Value = 1; }
        }
    }


}
