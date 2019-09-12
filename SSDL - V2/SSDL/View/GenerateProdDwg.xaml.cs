using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
//using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static SSDL.Publish;
using static SSDL.SSDLPlotingPublisher;

namespace SSDL.View
{
    /// <summary>
    /// Interaction logic for GenerateProdDwg.xaml
    /// </summary>
    public partial class GenerateProdDwg : Window, INotifyPropertyChanged
    {
        public string _pdfLocation;
        private SSDLPlotingSubscriber<double> _pbarMaximumSubscriber;
        private SSDLPlotingSubscriber<double> _pbarValueSubscriber;
        private SSDLPlotingSubscriber<string> _pbarStatusSubscriber;
        private SSDLPlotingSubscriber<string> _pbarHeaderSubscriber;
        public int maxAllowed = 30;
        public string remainingChars = " chars. remaining";
        public System.Collections.ObjectModel.ObservableCollection<clsTimeRollUp> _timeRollup;
        public System.Collections.ObjectModel.ObservableCollection<clsTimeRollUp> timeRollupCol
        {
            get { return _timeRollup; }
            set
            {
                _timeRollup = value;
                NotifyPropertyChanged();
            }
        }
        //= new ObservableCollection<clsTimeRollUp>();
        public bool readresult = false;
        public double totalTimeFromDwg = 0;
        public string drawingNumber = "";
        public double finalTimeFromUI = 0;

        public GenerateProdDwg(IntPtr ptHandler)
        {
            InitializeComponent();
            new WindowInteropHelper(this).Owner = ptHandler;
            InitializePublishers();
            // new WindowInteropHelper(this).Owner = ptHandler;

            _pbarMaximumSubscriber = new SSDLPlotingSubscriber<double>(SSDLPlotingPublisher.ProgressBarMaximumPublisher);
            _pbarMaximumSubscriber.Publisher.DataPublisher += Publisher_MaximumDataPublisher;

            _pbarValueSubscriber = new SSDLPlotingSubscriber<double>(SSDLPlotingPublisher.ProgressValuePublisher);
            _pbarValueSubscriber.Publisher.DataPublisher += Publisher_ValueDataPublisher;

            _pbarStatusSubscriber = new SSDLPlotingSubscriber<string>(SSDLPlotingPublisher.ProgressStatusPublisher);
            _pbarStatusSubscriber.Publisher.DataPublisher += Publisher_StatusDataPublisher;

            _pbarHeaderSubscriber = new SSDLPlotingSubscriber<string>(SSDLPlotingPublisher.ProgressHeaderPublisher);
            _pbarHeaderSubscriber.Publisher.DataPublisher += Publisher_HeaderDataPublisher;
        }

        #region Time Roll up
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
            calcPages();
        }

        private void calcTotal()
        {
            double _tempTotal = totalTimeFromDwg + timeRollupCol.Where(a => a.include == true).Sum(a => a.Time);
            _tempTotal = Math.Round(_tempTotal / 60, 2);//15-05-2018, added to change this to mins
            //TotalMins.Content = "Total : " + _tempTotal + " secs";//15-05-2018
            if (_tempTotal == 1)
                TotalMins.Content = "Total : " + _tempTotal + " min";
            else
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

        private void ckIncludedbHeader_Checked(object sender, RoutedEventArgs e)
        {

            // timeRollupCol.Select(a => a.include = true);
            timeRollupCol.All(c => { c.include = true; return true; });
            //dgvTimeRollup.ItemsSource = timeRollupCol;
            calcTotal();
        }

        private void ckIncludedbHeader_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ckIncludedbHeader_Unchecked(object sender, RoutedEventArgs e)
        {
            //timeRollupCol.Select(a => a.include = false);
            timeRollupCol.All(c => { c.include = false; return true; });
            //dgvTimeRollup.ItemsSource = timeRollupCol;

            calcTotal();
        }
        #endregion

        #region Layer
        private ObservableCollection<Layerlist> _layerlisodwg;

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public ObservableCollection<Layerlist> Layerlisodwg
        {
            get
            {
                return _layerlisodwg;
            }

            set
            {
                _layerlisodwg = value;
                NotifyPropertyChanged();
            }
        }
        private void ckIncludedPageHeader_Checked(object sender, RoutedEventArgs e)
        {
            Layerlisodwg.All(c => { c.Include = true; return true; });
            calcPages();
        }

        private void ckIncludedPageHeader_Unchecked(object sender, RoutedEventArgs e)
        {
            Layerlisodwg.All(c => { c.Include = false; return true; });
            calcPages();
        }

        private void ckIncludedPageHeader_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        private void InitializePublishers()
        {
            ProgressBarMaximumPublisher = new Publisher<double>();
            ProgressValuePublisher = new Publisher<double>();
            ProgressStatusPublisher = new Publisher<string>();
            ProgressHeaderPublisher = new Publisher<string>();
        }

        private void OnOrOffParticularLayer(bool isOff, string layerName, bool Page_1Available)
        {
            Document docPublish = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            try
            {
                using (DocumentLock docLock = docPublish.LockDocument())
                {
                    //Switch of and Thaw the layers that endswith "GRIDREF"
                    using (Transaction acTrans = docPublish.Database.TransactionManager.StartTransaction())
                    {
                        // Open the Layer table for read
                        LayerTable acLyrTbl;
                        acLyrTbl = acTrans.GetObject(docPublish.Database.LayerTableId, OpenMode.ForWrite) as LayerTable;

                        foreach (ObjectId acObjId in acLyrTbl)
                        {
                            LayerTableRecord acLyrTblRec;
                            acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForWrite) as LayerTableRecord;
                            if (acLyrTblRec.Name.ToString().Equals(layerName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                //if (acLyrTblRec.IsOff)
                                acLyrTblRec.IsOff = isOff ? false : true;
                                //if (acLyrTblRec.IsFrozen)
                                if (acLyrTblRec.IsFrozen)
                                    acLyrTblRec.IsFrozen = false;
                                try
                                {
                                    if (!acLyrTblRec.IsOff)
                                        acLyrTblRec.IsFrozen = false;
                                    // acLyrTblRec.IsFrozen = false;
                                    // if (!acLyrTblRec.IsFrozen)
                                    //   docPublish.Database.Clayer = acLyrTbl[acLyrTblRec.Name];
                                }
                                catch
                                {

                                }

                            }
                            else if (acLyrTblRec.Name.ToString().Equals("TITLE_2", StringComparison.InvariantCultureIgnoreCase) && layerName.Equals("PAGE_1", StringComparison.InvariantCultureIgnoreCase))
                            {
                                acLyrTblRec.IsOff = false;
                                if (acLyrTblRec.IsFrozen)
                                    acLyrTblRec.IsFrozen = false;
                            }
                            else if (acLyrTblRec.Name.ToString().Equals("TITLE_1", StringComparison.InvariantCultureIgnoreCase) || acLyrTblRec.Name.ToString().Equals("BORDER", StringComparison.InvariantCultureIgnoreCase))
                            {
                                acLyrTblRec.IsOff = false;
                                if (acLyrTblRec.IsFrozen)
                                    acLyrTblRec.IsFrozen = false;
                                try
                                {
                                    if (!acLyrTblRec.IsFrozen)
                                        docPublish.Database.Clayer = acLyrTbl[acLyrTblRec.Name];
                                }
                                catch { }
                            }
                            else
                            {
                                acLyrTblRec.IsOff = isOff;
                            }

                        }
                        acTrans.Commit();
                    }
                }
                //docPublish.CloseAndSave(docPublish.Name);
            }
            catch (Exception ex)
            {

            }
        }

        private void OffAllLayers(bool isOff)
        {
            Document docPublish = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            try
            {
                using (DocumentLock docLock = docPublish.LockDocument())
                {
                    //Switch of and Thaw the layers that endswith "GRIDREF"
                    using (Transaction acTrans = docPublish.Database.TransactionManager.StartTransaction())
                    {
                        // Open the Layer table for read
                        LayerTable acLyrTbl;
                        acLyrTbl = acTrans.GetObject(docPublish.Database.LayerTableId, OpenMode.ForWrite) as LayerTable;

                        foreach (ObjectId acObjId in acLyrTbl)
                        {
                            LayerTableRecord acLyrTblRec;
                            acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForWrite) as LayerTableRecord;

                            if (acLyrTblRec.Name.Equals(AcadFunctions.ActiveLayerName, StringComparison.InvariantCultureIgnoreCase))
                            {
                                docPublish.Database.Clayer = acLyrTbl[acLyrTblRec.Name];
                            }

                            if (AcadFunctions.ActiveLayers.Contains(acLyrTblRec.Name))
                            {
                                acLyrTblRec.IsOff = false;
                            }
                            else
                            {
                                acLyrTblRec.IsOff = isOff;
                            }
                        }
                        acTrans.Commit();
                    }
                }
                //docPublish.CloseAndSave(docPublish.Name);
                //Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CloseAll();
            }
            catch (Exception ex)
            {

            }
        }

        private void InitializeLayouts()
        {
            Document docPublish = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            ArrayList layoutIdList = GetLayoutIdList(docPublish.Database);
            using (docPublish.LockDocument())
            {
                using (Transaction trans = docPublish.TransactionManager.StartTransaction())
                {
                    LayoutManager acLayoutMgr = LayoutManager.Current;
                    Layout _layout = null;
                    foreach (ObjectId layoutId in layoutIdList)
                    {
                        _layout = trans.GetObject(layoutId, OpenMode.ForRead) as Layout;
                        acLayoutMgr.CurrentLayout = _layout.LayoutName;
                    }
                    acLayoutMgr.CurrentLayout = "Model";
                    trans.Commit();
                }
            }
            //docPublish.CloseAndSave(docPublish.Name);
        }
        public static string PadNumbers(string input)
        {
            return Regex.Replace(input, "[0-9]+", match => match.Value.PadLeft(10, '0'));
        }

        private void Publisher_HeaderDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<string> e)
        {
            // lblHeader.Dispatcher.Invoke(() => { lblHeader.Content = e.Message; });
            DoEventsHandler.DoEvents();
        }

        private void Publisher_StatusDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<string> e)
        {
            //lblStatusPublisher.Dispatcher.Invoke(() => { lblStatusPublisher.Content = e.Message; });
            DoEventsHandler.DoEvents();
        }

        private void Publisher_ValueDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<double> e)
        {
            //pbStatus.Dispatcher.Invoke(() => { pbStatus.Value = pbStatus.Value + e.Message; });
            //txtPercentage.Dispatcher.Invoke(() =>
            //{
            //    var value = Convert.ToInt32((pbStatus.Value / pbStatus.Maximum) * 100);
            //    txtPercentage.Text = value + "%";
            //});
            DoEventsHandler.DoEvents();
        }

        private void Publisher_MaximumDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<double> e)
        {
            //pbStatus.Dispatcher.Invoke(() => { pbStatus.Value = 0; });
            //pbStatus.Dispatcher.Invoke(() => { pbStatus.Maximum = e.Message; });
            DoEventsHandler.DoEvents();
        }

        public static ArrayList GetLayoutIdList(Database db)
        {
            ArrayList layoutList = new ArrayList();

            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm =
                db.TransactionManager;

            using (Transaction myT = tm.StartTransaction())
            {
                DBDictionary dic = (DBDictionary)tm.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, false);
                DbDictionaryEnumerator index = dic.GetEnumerator();

                while (index.MoveNext())
                {
                    Layout acLayout = tm.GetObject(index.Current.Value,
                                            OpenMode.ForRead) as Layout;
                    if (acLayout.LayoutName.Equals("model", StringComparison.InvariantCultureIgnoreCase) ||
                          acLayout.LayoutName.Equals("modal", StringComparison.InvariantCultureIgnoreCase))
                        layoutList.Add(index.Current.Value);
                }
                myT.Commit();
            }

            return layoutList;
        }

        private void SetDefaultLayout(Document docPublish)
        {
            using (docPublish.LockDocument())
            {
                using (Transaction trans = docPublish.TransactionManager.StartTransaction())
                {
                    LayoutManager acLayoutMgr = LayoutManager.Current;
                    acLayoutMgr.CurrentLayout = "Model";
                    trans.Commit();
                }
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void calcPages()
        {
            lblTotalPages.Content = "Total Pages: " + Layerlisodwg.Where(a => a.Include == true).Count();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            Layerlisodwg = new System.Collections.ObjectModel.ObservableCollection<Layerlist>(AcadFunctions.GetLayerlist());



            System.Data.DataTable dt = DatabaseClass.GetTimeRollupCategories();

            List<View.clsTimeRollUp> _col = new List<View.clsTimeRollUp>();
            foreach (DataRow dr in dt.Rows)
            {
                int time = 0;
                if (dr["fldTimeAllocated"] != null)
                {
                    if (!string.IsNullOrEmpty(dr["fldTimeAllocated"].ToString()))
                        time = Convert.ToInt32(dr["fldTimeAllocated"].ToString());
                }
                _col.Add(new View.clsTimeRollUp() { include = false, Category = dr["fldCategory"].ToString(), Time = time });
            }
            System.Collections.ObjectModel.ObservableCollection<View.clsTimeRollUp> myCollection = new System.Collections.ObjectModel.ObservableCollection<View.clsTimeRollUp>(_col as List<View.clsTimeRollUp>);
            timeRollupCol = myCollection;

            textBoxdesc.MaxLength = maxAllowed;
            updateRemainingChars();
            dgvTimeRollup.ItemsSource = timeRollupCol;
            //---------------Added by Sundari on 1-11-2018
            BOM od = new BOM();
            System.Data.DataTable dtnew = od.Get_Drawingdata(updateDwgAttr: false);
            System.Data.DataTable dtnewExcel = new System.Data.DataTable();
            if (dtnew.Rows.Count > 0)
            {
                dtnewExcel = dtnew.DefaultView.ToTable(false, "Partno", "Description", "Timerollup", "SurfaceArea");
            }
            string tempSum = (dtnewExcel.Rows.Count > 0) ? dtnewExcel.AsEnumerable().Sum(x => Convert.ToDouble(x["Timerollup"])).ToString() : "0";
            totalTimeFromDwg = Convert.ToDouble(tempSum);
            //--------------------------
            calcTotal();

            dgvLayerlist.ItemsSource = Layerlisodwg.Where(x => x.LayerName.StartsWith("PAGE", StringComparison.InvariantCultureIgnoreCase)).ToList();
            dgvLayerlist.Items.Refresh();
        }

        private void btnSelectallTimeRollup_Click(object sender, RoutedEventArgs e)
        {
            ckIncludedbHeader_Checked(null, null);
        }

        private void btnUnSelectAllTimeRolluo_Click(object sender, RoutedEventArgs e)
        {
            ckIncludedbHeader_Unchecked(null, null);
        }

        private void btnUnSelectAllLayer_Click(object sender, RoutedEventArgs e)
        {
            ckIncludedPageHeader_Unchecked(null, null);
        }

        private void btnSelectallLayer_Click(object sender, RoutedEventArgs e)
        {
            ckIncludedPageHeader_Checked(null, null);
        }

        private BackgroundWorker worker = new BackgroundWorker();
        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            _pdfLocation = getpath.IniReadValue("FilePath", "PDFOUTPUT_PATH");
            if (textBoxdesc.Text.Trim().Length <= 0 || string.IsNullOrEmpty(textBoxdesc.Text.Trim()))
            {
                MessageBox.Show("Please specify Hospital Title");
                return;
            }
            else if (Layerlisodwg.Where(a => a.Include == true).Count() <= 0)
            {
                MessageBox.Show("Select minimum 1 Page");
                return;
            }
            //else if (timeRollupCol.Where(a => a.include == true).Count() <= 0)
            //{
            //    MessageBox.Show("Select minimum 1 Category");
            //    return;
            //}
            else
            {


                
                btnGenerate.IsEnabled = false;
                implementFn();
                btnGenerate.IsEnabled = true;
                //worker.DoWork += Worker_DoWork;
                //worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
                //worker.ProgressChanged += Worker_ProgressChanged;
                //worker.RunWorkerAsync();
            }
        }
        private void implementFn()
        {
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string path = getpath.IniReadValue("FilePath", "BOM_OUTPUT_LOCATION");
            string BOMExcelfulltitle = System.IO.Path.Combine(path, System.IO.Path.GetFileNameWithoutExtension(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name) + "_BOM.xlsx");

            //string BOMExcelfulltitle = System.IO.Path.Combine(path, System.IO.Path.GetFileNameWithoutExtension(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name) + "_"+DateTime.Now.ToString("yyyy_MM_dd")+"_XML.xlsx");

            //Generate button functionality
            //Create BOM Excel
            try
            {
                statusText = "Generating Production Drawing…";
                lblpbBarContent.Content = statusText;
                pbBar.Value = (5);
                DoEventsHandler.DoEvents();
                statusText = "Re-populating Drawing Item Numbers... (Task 1 of 6)";
                lblpbBarContent.Content = statusText;
                pbBar.Value = (8);
                DoEventsHandler.DoEvents();
                statusText = "Generating BOM Spreadsheet... (Task 2 of 6)";
                lblpbBarContent.Content = statusText;
                pbBar.Value = 10;
                DoEventsHandler.DoEvents();
                BOM od = new BOM();
                od.strBomName = textBoxdesc.Text;
                od.Generate(ExcelOutput: true);
                //statusText = "Generating BOM Spreadsheet... (Task 2 of 6)";
                //lblpbBarContent.Content = statusText;
                ////pbBar.Value = (40);
                //pbBar.Value = pbBar.Value + 10;
                //DoEventsHandler.DoEvents();
                if (od.sheetvalue != null && od.sheetvalue.Count > 0)
                {
                    String title = "";
                    foreach (KeyValuePair<string, string> kvp in od.sheetvalue)
                    {
                        if (kvp.Key.Equals("DRG", StringComparison.InvariantCultureIgnoreCase))
                        {
                            title = kvp.Value;
                            break;
                        }
                    }
                    BOMExcelfulltitle = System.IO.Path.Combine(path, title + "_BOM.xlsx");
                }
            }
            catch { }
            //statusText = "Generating BOM Spreadsheet... (Task 2 of 6)";
            //lblpbBarContent.Content = statusText;
            ////pbBar.Value = (40);
            //pbBar.Value = pbBar.Value + 10;
            //DoEventsHandler.DoEvents();
            statusText = "Generating Time Rollup Spreadsheet... (Task 3 of 6)";
            lblpbBarContent.Content = statusText;
            //pbBar.Value = (60);
            pbBar.Value = pbBar.Value + 30;
            DoEventsHandler.DoEvents();
            try
            {
                timeRollupExcelandXML();
            }
            catch { }
            statusText = "Generating Time Rollup XML... (Task 4 of 6)";
            lblpbBarContent.Content = statusText;
            //pbBar.Value = (70);
            pbBar.Value = pbBar.Value + 15;
            DoEventsHandler.DoEvents();
            //Create Drawing PDF with BOM Excel added to Bottom
            statusText = "Generating Drawing PDF... (Task 5 of 6)";
            lblpbBarContent.Content = statusText;            
            pbBar.Value = pbBar.Value + (20);
            DoEventsHandler.DoEvents();
            try
            {
                generatePDF(BOMExcelfulltitle);//Adding Excel to PDF is pending
            }
            catch { }
            //statusText = "Appending BOM Spreadsheet to PDF... (Task 6 of 6)";
            //lblpbBarContent.Content = statusText;
            ////pbBar.Value = (90);
            //pbBar.Value = pbBar.Value + 10;
            //DoEventsHandler.DoEvents();
            statusText = "All tasks have Completed!";
            lblpbBarContent.Content = statusText;            
            pbBar.Value = (100);
            DoEventsHandler.DoEvents();
        }


        string statusText = "";
        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblpbBarContent.Content = statusText;
            pbBar.Value = e.ProgressPercentage;

            if (e.ProgressPercentage == 100)
            {
                btnGenerate.IsEnabled = true;                
            }
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //Update UI
            btnGenerate.IsEnabled = true;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string path = getpath.IniReadValue("FilePath", "BOM_OUTPUT_LOCATION");
            string BOMExcelfulltitle = System.IO.Path.Combine(path, System.IO.Path.GetFileNameWithoutExtension(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name) + "_BOM.xlsx");


            //Generate button functionality
            //Create BOM Excel
            try
            {
                statusText = "Re-populating Drawing Item Numbers... (Task 1 of 6)";
                worker.ReportProgress(10);
                BOM od = new BOM();
                od.Generate(ExcelOutput: true);
               
            }
            catch { }
            statusText = "Generating BOM Spreadsheet... (Task 2 of 6)";
            worker.ReportProgress(40);

            //Create Time Rollup Excel
            //Create Time Rollup XML
            statusText = "Generating Time Rollup Spreadsheet... (Task 3 of 6)";
            worker.ReportProgress(60);
            try
            {
                timeRollupExcelandXML();
            }
            catch { }
            statusText = "Generating Time Rollup XML... (Task 4 of 6)";
            worker.ReportProgress(70);
            //Create Drawing PDF with BOM Excel added to Bottom
            statusText = "Generating Drawing PDF... (Task 5 of 6)";
            worker.ReportProgress(80);
            try
            {
                generatePDF(BOMExcelfulltitle);//Adding Excel to PDF is pending
            }
            catch { }
            statusText = "Appending BOM Spreadsheet to PDF... (Task 5 of 6)";
            worker.ReportProgress(90);

            statusText = "All tasks have Completed!";
            worker.ReportProgress(100);
        }

        #region Generate PDF

        private void generatePDF(string BOMExcel)
        {
            string docName = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name;
            try
            {
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                acDoc.Database.SaveAs(acDoc.Name, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
            }
            catch (Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(ex.ToString());
            }
            if (AcadFunctions.AcadLayersList.Count > 0)
            {
                if (Layerlisodwg.Any(x => x.Include))
                {
                    InitializeLayouts();
                    Document docPublish = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    var lstLayers = Layerlisodwg.Where(x => x.Include && x.EntityAvail).Select(x => x).ToList();


                    ProgressHeaderPublisher.PublishData("Processing...");
                    ProgressStatusPublisher.PublishData("");
                    ProgressBarMaximumPublisher.PublishData(lstLayers.Count + 1);

                    int cnt = 0;

                    bool Page_1Available = false;
                    if (lstLayers.ToList().Count > 0)
                    {
                        Page_1Available = lstLayers.Where(a => a.Include == true && (a.LayerName.Equals("PAGE_1", StringComparison.InvariantCultureIgnoreCase) || a.LayerName.Equals("PAGE1", StringComparison.InvariantCultureIgnoreCase))).Any();
                    }
                    lstLayers.ToList().ForEach(x =>
                    {
                        cnt++;
                        ProgressHeaderPublisher.PublishData("Processing " + cnt + " of " + lstLayers.Count);
                        ProgressStatusPublisher.PublishData("Publishing " + x.LayerName + "...");
                        if (x.Include)
                        {
                            var layerName = x.LayerName;
                            OnOrOffParticularLayer(true, layerName, Page_1Available);

                            docPublish.SendStringToExecute("zoom e ", true, false, true);

                            Thread.Sleep(1000);

                            new SingleSheetPdf(_pdfLocation, GetLayoutIdList(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database)).PublishNew(_pdfLocation, x.LayerName);

                            OnOrOffParticularLayer(false, layerName, Page_1Available);
                        }
                        ProgressValuePublisher.PublishData(1);
                    });
                    SetDefaultLayout(docPublish);
                    OffAllLayers(true);

                    ProgressHeaderPublisher.PublishData("Processing....");
                    ProgressStatusPublisher.PublishData("Merging pdf");

                    try
                    {
                        string path = _pdfLocation;
                        string[] pdfFiles = Directory.GetFiles(path, "*.pdf").Select(f => new FileInfo(f)).OrderBy(f => f.CreationTime).Select(d => d.FullName).ToArray();
                        var completedFiles = new List<string>();
                        docName = System.IO.Path.GetFileNameWithoutExtension(docName);
                        foreach (var item in pdfFiles)
                        {
                            if (System.IO.Path.GetFileNameWithoutExtension(item).StartsWith(docName))
                            {
                                var layer = Layerlisodwg.Where(y => y.Include && item.Contains(y.LayerName)).Select(y => y).FirstOrDefault();
                                if (layer != null)
                                {
                                    completedFiles.Add(item);
                                }
                            }
                        }
                        statusText = "Appending BOM Spreadsheet to PDF... (Task 6 of 6)";
                        lblpbBarContent.Content = statusText;
                        pbBar.Value = pbBar.Value + 10;
                        DoEventsHandler.DoEvents();
                        // step 1: creation of a document-object
                        iTextSharp.text.Document document = new iTextSharp.text.Document();

                        if (File.Exists(System.IO.Path.Combine(path, docName + ".pdf")))
                            File.Delete(System.IO.Path.Combine(path, docName + ".pdf"));
                        if (completedFiles.Count > 0)
                        {
                            // step 2: we create a writer that listens to the document
                            PdfCopy writer = new PdfCopy(document, new FileStream(System.IO.Path.Combine(path, docName + ".pdf"), FileMode.Create));
                            if (writer == null)
                            {
                                return;
                            }

                            // step 3: we open the document
                            document.Open();

                            foreach (string fileName in completedFiles)
                            {
                                // we create a reader for a certain document
                                PdfReader reader = new PdfReader(fileName);
                                reader.ConsolidateNamedDestinations();

                                // step 4: we add content
                                for (int i = 1; i <= reader.NumberOfPages; i++)
                                {
                                    PdfImportedPage page = writer.GetImportedPage(reader, i);
                                    writer.AddPage(page);
                                }

                                PRAcroForm form = reader.AcroForm;
                                if (form != null)
                                {
                                    writer.CopyAcroForm(reader);
                                }

                                reader.Close();
                            }

                            // step 5: we close the document and writer
                            writer.Close();
                            document.Close();
                            if (System.IO.File.Exists(BOMExcel))
                            {
                                String strDwgPDF = System.IO.Path.Combine(path, docName + ".pdf");
                                string strBOMPDF = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(BOMExcel), System.IO.Path.GetFileNameWithoutExtension(BOMExcel) + ".pdf");
                                string strResultPDF = System.IO.Path.Combine(path, docName + "_PROD.pdf");
                                string shellCommand = "C:\\ACAD2018_LivaNova\\Interface\\PDFTK.exe " + strDwgPDF + " " + strBOMPDF + " cat output " + strResultPDF;
                                AcadCommands oAcad = new AcadCommands();
                                oAcad.ExecuteCommandSync(shellCommand);
                            }

                            foreach (var item in completedFiles)
                            {
                                try
                                {
                                    if (System.IO.File.Exists(item))
                                        System.IO.File.Delete(item);
                                }
                                catch { }
                            }
                            ProgressValuePublisher.PublishData(1);
                            ProgressHeaderPublisher.PublishData("Processed " + cnt + " of " + lstLayers.Count);
                            ProgressStatusPublisher.PublishData("Completed");
                            //MessageBox.Show("PDF Generation Completed", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                            //10-05-2018
                            if (System.IO.File.Exists(System.IO.Path.Combine(path, docName + "_PROD.pdf")))
                            {
                                System.Diagnostics.Process.Start(System.IO.Path.Combine(path, docName + "_PROD.pdf"));
                                //System.IO.File.Open(System.IO.Path.Combine(path, docName + ".pdf"), FileMode.Open);
                            }
                        }
                        else
                        {

                            ProgressValuePublisher.PublishData(1);
                            ProgressHeaderPublisher.PublishData("Processed " + cnt + " of " + lstLayers.Count);
                            ProgressStatusPublisher.PublishData("Failed");
                            MessageBox.Show("Failed to merge pdf.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
                else
                {
                    MessageBox.Show("Atleast one layer should be selected.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                }

            }
        }

        #endregion

        #region Time Roll up Excel and xml



        private void timeRollupExcelandXML()
        {
            BOM od = new BOM();

            System.Data.DataTable dtnew = od.Get_Drawingdata(updateDwgAttr:false);

            if (dtnew == null)
            {
                System.Windows.MessageBox.Show("No items found in drawing");
                return;
            }
            else if (dtnew.Rows.Count <= 0)
            {
                System.Windows.MessageBox.Show("No items found in drawing");
                return;
            }
            od.timeRollupCol = timeRollupCol;
            string sumresult = "", drgnumber = "", SurfaceAreaSum = "";
            string projectDesc = textBoxdesc.Text;// od.getProjectDesc();//20-Oct-2018
            bool result = od.timerollupfunctionality(out sumresult, out drgnumber, out SurfaceAreaSum, out projectDesc, true, readDesc: false, divideBy60: false, updateDWGAttr: false,prevDesc: textBoxdesc.Text,strTotalTime: finalTimeFromUI.ToString()); //02-05-2018

            od.SendXml(false, projectDesc, sumresult, drgnumber, SurfaceAreaSum);


        }

        #endregion
    }
}
