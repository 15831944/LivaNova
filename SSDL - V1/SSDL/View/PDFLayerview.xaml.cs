using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static SSDL.Publish;
using static SSDL.SSDLPlotingPublisher;

namespace SSDL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class PDFLayerview : INotifyPropertyChanged
    {
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

        // public System.Collections.ObjectModel.ObservableCollection<Layerlist> _layerlisodwg;
        //public System.Collections.ObjectModel.ObservableCollection<Layerlist> Layerlisodwg
        //{
        //    get { return _layerlisodwg; }
        //    set
        //    {
        //        _layerlisodwg = value;
        //    }
        //}
        private string _pdfLocation = string.Empty;

        private SSDLPlotingSubscriber<double> _pbarMaximumSubscriber;
        private SSDLPlotingSubscriber<double> _pbarValueSubscriber;
        private SSDLPlotingSubscriber<string> _pbarStatusSubscriber;
        private SSDLPlotingSubscriber<string> _pbarHeaderSubscriber;

        public PDFLayerview(IntPtr ptHandler)
        {
            InitializeComponent();
            InitializePublishers();
            new WindowInteropHelper(this).Owner = ptHandler;

            _pbarMaximumSubscriber = new SSDLPlotingSubscriber<double>(SSDLPlotingPublisher.ProgressBarMaximumPublisher);
            _pbarMaximumSubscriber.Publisher.DataPublisher += Publisher_MaximumDataPublisher;

            _pbarValueSubscriber = new SSDLPlotingSubscriber<double>(SSDLPlotingPublisher.ProgressValuePublisher);
            _pbarValueSubscriber.Publisher.DataPublisher += Publisher_ValueDataPublisher;

            _pbarStatusSubscriber = new SSDLPlotingSubscriber<string>(SSDLPlotingPublisher.ProgressStatusPublisher);
            _pbarStatusSubscriber.Publisher.DataPublisher += Publisher_StatusDataPublisher;

            _pbarHeaderSubscriber = new SSDLPlotingSubscriber<string>(SSDLPlotingPublisher.ProgressHeaderPublisher);
            _pbarHeaderSubscriber.Publisher.DataPublisher += Publisher_HeaderDataPublisher;

            Layerlisodwg = new System.Collections.ObjectModel.ObservableCollection<Layerlist>(AcadFunctions.GetLayerlist());
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            _pdfLocation = getpath.IniReadValue("FilePath", "PDFOUTPUT_PATH");

        }

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

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{//17-July-2018
            //    string[] existingPdfFiles = Directory.GetFiles(_pdfLocation, "*.pdf");
            //    if (existingPdfFiles != null)
            //        if (existingPdfFiles.Count() > 0)
            //        {
            //            MessageBox.Show("Please delete all existing files in "+ _pdfLocation +" before generating PDF", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            //            return;
            //        }
            //}
            //catch { }
            string docName = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name;
            try
            {//10-05-2018
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                acDoc.Database.SaveAs(acDoc.Name, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
                //Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.Save();
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
                    //Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(docName);
                    Document docPublish = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    //Layerlisodwg.Where(x => x.LayerName.Equals("TITLE_1", StringComparison.InvariantCultureIgnoreCase) ||
                    // x.LayerName.Equals("BORDER", StringComparison.InvariantCultureIgnoreCase)).ToList().ForEach(a => a.Include = true);

                    var lstLayers = Layerlisodwg.Where(x => x.Include && x.EntityAvail).Select(x => x).ToList();


                    gridProcess.Visibility = System.Windows.Visibility.Visible;
                    ProgressHeaderPublisher.PublishData("Processing...");
                    txtPercentage.Text = string.Empty;
                    ProgressStatusPublisher.PublishData("");
                    ProgressBarMaximumPublisher.PublishData(lstLayers.Count + 1);
                    pbStatus.Visibility = System.Windows.Visibility.Visible;
                    txtPercentage.Visibility = System.Windows.Visibility.Visible;

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
                        if (x.Include)// && x.EntityAvail)
                        {
                            var layerName = x.LayerName;
                            OnOrOffParticularLayer(true, layerName, Page_1Available);

                            docPublish.SendStringToExecute("zoom e ", true, false, true);

                            Thread.Sleep(1000);

                            //new SingleSheetPdf(_pdfLocation, GetLayoutIdList(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database)).Publish(x.LayerName);

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
                        string[] pdfFiles = Directory.GetFiles(path, "*.pdf").Select(f => new FileInfo(f)).OrderBy(f => f.CreationTime).Select(d=>d.FullName).ToArray();
                       // pdfFiles = pdfFiles.OrderBy(z=> PadNumbers(System.IO.Path.GetFileNameWithoutExtension(z)))
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
                            MessageBox.Show("PDF Generation Completed", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                            gridProcess.Visibility = System.Windows.Visibility.Collapsed;
                            pbStatus.Visibility = System.Windows.Visibility.Collapsed;
                            txtPercentage.Visibility = System.Windows.Visibility.Collapsed;
                            //10-05-2018
                            if (System.IO.File.Exists(System.IO.Path.Combine(path, docName + ".pdf")))
                            {
                                System.Diagnostics.Process.Start(System.IO.Path.Combine(path, docName + ".pdf"));
                                //System.IO.File.Open(System.IO.Path.Combine(path, docName + ".pdf"), FileMode.Open);
                            }
                        }
                        else
                        {

                            ProgressValuePublisher.PublishData(1);
                            ProgressHeaderPublisher.PublishData("Processed " + cnt + " of " + lstLayers.Count);
                            ProgressStatusPublisher.PublishData("Failed");
                            MessageBox.Show("Failed to merge pdf.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                            gridProcess.Visibility = System.Windows.Visibility.Collapsed;
                            pbStatus.Visibility = System.Windows.Visibility.Collapsed;
                            txtPercentage.Visibility = System.Windows.Visibility.Collapsed;
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
        public static string PadNumbers(string input)
        {
            return Regex.Replace(input, "[0-9]+", match => match.Value.PadLeft(10, '0'));
        }

        private void Publisher_HeaderDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<string> e)
        {
            lblHeader.Dispatcher.Invoke(() => { lblHeader.Content = e.Message; });
            DoEventsHandler.DoEvents();
        }

        private void Publisher_StatusDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<string> e)
        {
            lblStatusPublisher.Dispatcher.Invoke(() => { lblStatusPublisher.Content = e.Message; });
            DoEventsHandler.DoEvents();
        }

        private void Publisher_ValueDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<double> e)
        {
            pbStatus.Dispatcher.Invoke(() => { pbStatus.Value = pbStatus.Value + e.Message; });
            txtPercentage.Dispatcher.Invoke(() =>
            {
                var value = Convert.ToInt32((pbStatus.Value / pbStatus.Maximum) * 100);
                txtPercentage.Text = value + "%";
            });
            DoEventsHandler.DoEvents();
        }

        private void Publisher_MaximumDataPublisher(object sender, SSDLPlotingPublisher.MessageArgument<double> e)
        {
            pbStatus.Dispatcher.Invoke(() => { pbStatus.Value = 0; });
            pbStatus.Dispatcher.Invoke(() => { pbStatus.Maximum = e.Message; });
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //Layerlisodwg = new System.Collections.ObjectModel.ObservableCollection<Layerlist>(AcadFunctions.GetLayerlist());
            dgvLayerlist.ItemsSource = Layerlisodwg.Where(x => x.LayerName.StartsWith("PAGE", StringComparison.InvariantCultureIgnoreCase)).ToList();
                                            //commented on 30-Aug-2018, to display only layers starts with PAGE
                                            //Layerlisodwg.Where(x => !x.LayerName.Equals("TITLE_1", StringComparison.InvariantCultureIgnoreCase) &&                                            !x.LayerName.Equals("BORDER", StringComparison.InvariantCultureIgnoreCase)).ToList();
                                            //dgvLayerlist.ItemsSource = Layerlisodwg;
            dgvLayerlist.Items.Refresh();

            //dgvLayerlist.Items.SortDescriptions.Add(new SortDescription("LayerName", ListSortDirection.Ascending));
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class Layerlist : INotifyPropertyChanged
    {
        private bool _include;
        private string _layerName;
        private bool _entityAvail;





        public bool Include
        {
            get
            {
                return _include;
            }

            set
            {
                _include = value;
                NotifyPropertyChanged();
            }
        }

        public string LayerName
        {
            get
            {
                return _layerName;
            }

            set
            {
                _layerName = value;
                NotifyPropertyChanged();
            }
        }

        public bool EntityAvail
        {
            get
            {
                return _entityAvail;
            }

            set
            {
                _entityAvail = value;
                NotifyPropertyChanged();
            }
        }

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
