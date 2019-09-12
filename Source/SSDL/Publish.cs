using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Publishing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SSDL
{
    public class Publish
    {
        public Publish()
        { }

        // Base class for the different configurations
        public abstract class PlotToFileConfig
        {
            // Private fields
            private string dsdFile, dwgFile, outputDir, outputFile, plotType;
            private int sheetNum;
            private ArrayList layouts;
            private const string LOG = "publish.log";

            // Base constructor
            public PlotToFileConfig(string outputDir, ArrayList layouts, string plotType)
            {
                Database db = HostApplicationServices.WorkingDatabase;
                this.dwgFile = db.Filename;
                this.outputDir = outputDir;
                this.dsdFile = Path.ChangeExtension(this.dwgFile, "dsd");
                this.layouts = layouts;
                this.plotType = plotType;
                string ext = plotType == "0" || plotType == "1" ? "dwf" : "pdf";
                this.outputFile = Path.Combine(
                    this.outputDir,
                    Path.ChangeExtension(Path.GetFileName(this.dwgFile), ext));
            }


            public void PublishNew(string pdfLocation, string layerName)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication.GetType().InvokeMember("ZoomExtents", BindingFlags.InvokeMethod, null, Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, null);
                Database db = doc.Database;
                Editor ed = doc.Editor;

                string dwgFileName =
                    Application.GetSystemVariable("DWGNAME") as string;
                string dwgPath =
                    Application.GetSystemVariable("DWGPREFIX") as string;

                string name =
                    System.IO.Path.GetFileNameWithoutExtension(dwgFileName);

                //find a temp location.
                string strTemp = System.IO.Path.GetTempPath();

                //PromptStringOptions options =
                //    new PromptStringOptions("Specific the DWF file name");
                //options.DefaultValue = "c:\\temp\\" + name + ".dwf";
                //PromptResult result = ed.GetString(options);

                //if (result.Status != PromptStatus.OK)
                //    return;

                //get the layout ObjectId List
                System.Collections.ArrayList layoutList = null;
                try
                {
                    layoutList = layouts;
                }
                catch
                {
                    //  Application.ShowAlertDialog("Unable to get layouts name");
                    return;
                }

                Publisher publisher = Application.Publisher;

                //put the plot in foreground
                short bgPlot =
                    (short)Application.GetSystemVariable("BACKGROUNDPLOT");
                Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                using (Transaction tr =
                            db.TransactionManager.StartTransaction())
                {
                    DsdEntryCollection collection =
                                    new DsdEntryCollection();

                    Layout _layout = null;
                    foreach (ObjectId layoutId in layoutList)
                    {
                        Layout layout =  tr.GetObject(layoutId, OpenMode.ForWrite) as Layout;

                        #region plot Extends

                        Point3d pMin = layout.Extents.MinPoint;
                        Point3d pMax = layout.Extents.MaxPoint;
                        PlotSettingsValidator psv = PlotSettingsValidator.Current;
                        PlotSettings ps = new PlotSettings(layout.ModelType);
                        ps.CopyFrom(layout);
                        psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                        bool bAreaIsLandscape = true;
                        if (layout.ModelType) bAreaIsLandscape = (db.Extmax.X - db.Extmin.X) > (db.Extmax.Y - db.Extmin.Y) ? true : false;
                        else bAreaIsLandscape = (layout.Extents.MaxPoint.X - layout.Extents.MinPoint.X) > (layout.Extents.MaxPoint.Y - layout.Extents.MinPoint.Y) ? true : false;

                        bool bPaperIsLandscape = (ps.PlotPaperSize.X > ps.PlotPaperSize.Y) ? true : false;
                        if (bPaperIsLandscape != bAreaIsLandscape) psv.SetPlotRotation(ps, PlotRotation.Degrees270);
                        ps.PlotSettingsName = layout.LayoutName + "_plot";
                        ps.AddToPlotSettingsDictionary(db);
                        tr.AddNewlyCreatedDBObject(ps, true);
                        psv.RefreshLists(ps);
                        layout.CopyFrom(ps);

                        #endregion

                        DsdEntry entry = new DsdEntry();

                        entry.DwgName = dwgPath + dwgFileName;
                        entry.Layout = layout.LayoutName;
                        entry.Title = Path.GetFileNameWithoutExtension(this.dwgFile) + "-" + layerName + "-" + layout.LayoutName;
                        entry.Nps = ps.PlotSettingsName;// "AA";
                        _layout = layout;
                        collection.Add(entry);
                    }

                    DsdData dsdData = new DsdData();

                    dsdData.SheetType = SheetType.SinglePdf;
                    dsdData.ProjectPath = pdfLocation;
                    dsdData.DestinationName = pdfLocation + name + " " + layerName + ".pdf";


                    dsdData.SetUnrecognizedData("PwdProtectPublishedDWF", "FALSE");
                    dsdData.SetUnrecognizedData("PromptForPwd", "FALSE");
                    dsdData.SetUnrecognizedData("INCLUDELAYER", "FALSE");
                    dsdData.NoOfCopies = 1;
                    dsdData.IsHomogeneous = false;

                    string dsdFile = pdfLocation + name + layerName + ".dsd";

                    if (System.IO.File.Exists(dsdFile))
                        System.IO.File.Delete(dsdFile);

                    dsdData.SetDsdEntryCollection(collection);
                    dsdData.PromptForDwfName = false;

                    //Workaround to avoid promp for dwf file name
                    //set PromptForDwfName=FALSE in
                    //dsdData using StreamReader/StreamWriter
                    dsdData.WriteDsd(dsdFile);

                    StreamReader sr = new StreamReader(dsdFile);
                    string str = sr.ReadToEnd();
                    sr.Close();

                    //str =
                    //    str.Replace("PromptForDwfName=TRUE",
                    //                    "PromptForDwfName=FALSE");
                    str =
                        str.Replace("IncludeLayer=TRUE",
                                               "IncludeLayer=FALSE");

                    StreamWriter sw = new StreamWriter(dsdFile);
                    sw.Write(str);
                    sw.Close();

                    dsdData.ReadDsd(dsdFile);

                    //18-07-2018
                    //PlotSettings ps = new PlotSettings(_layout.ModelType);
                    //ps.CopyFrom(_layout);

                    //PlotSettingsValidator psv = PlotSettingsValidator.Current;
                    //psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                    //psv.SetUseStandardScale(ps, true);
                    //psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                    //psv.SetPlotCentered(ps, true);


                    //// We'll use the standard DWF PC3, as

                    //// for today we're just plotting to file


                    //psv.SetPlotConfigurationName(ps, "DWG to PDF.PC3", "ANSI_A_(8.50_x_11.00_Inches)");

                    using (PlotConfig pc = PlotConfigManager.SetCurrentConfig("DWG to PDF.PC3"))
                    {
                        publisher.PublishExecute(dsdData, pc);
                    }

                    System.IO.File.Delete(dsdFile);

                    tr.Commit();
                }

                //reset the background plot value
                Application.SetSystemVariable("BACKGROUNDPLOT", bgPlot);
            }


            // Plot the layouts
            public void Publish(string layerName)
            {
                if (TryCreateDSD(layerName))
                {
                    object bgp = Application.GetSystemVariable("BACKGROUNDPLOT");
                    object ctab = Application.GetSystemVariable("CTAB");
                    try
                    {
                        Application.SetSystemVariable("BACKGROUNDPLOT", 0);


                        Publisher publisher = Application.Publisher;
                        PlotProgressDialog plotDlg = new PlotProgressDialog(false, this.sheetNum, true);
                        publisher.PublishDsd(this.dsdFile, plotDlg);
                        plotDlg.Destroy();
                        File.Delete(this.dsdFile);
                    }
                    catch (System.Exception exn)
                    {
                        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                        ed.WriteMessage("\nError: {0}\n{1}", exn.Message, exn.StackTrace);
                        throw;
                    }
                    finally
                    {
                        Application.SetSystemVariable("BACKGROUNDPLOT", bgp);
                        Application.SetSystemVariable("CTAB", ctab);
                    }
                }
            }

            // Creates the DSD file from a template (default options)
            private bool TryCreateDSD(string layerName)
            {
                using (DsdData dsd = new DsdData())
                using (DsdEntryCollection dsdEntries = CreateDsdEntryCollection(this.layouts, Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database, layerName))
                {
                    if (dsdEntries == null || dsdEntries.Count <= 0) return false;

                    if (!Directory.Exists(this.outputDir))
                        Directory.CreateDirectory(this.outputDir);

                    this.sheetNum = dsdEntries.Count;

                    dsd.SetDsdEntryCollection(dsdEntries);

                    dsd.SetUnrecognizedData("PwdProtectPublishedDWF", "FALSE");
                    dsd.SetUnrecognizedData("PromptForPwd", "FALSE");
                    dsd.SetUnrecognizedData("INCLUDELAYER", "FALSE");
                    dsd.NoOfCopies = 1;
                    dsd.DestinationName = this.outputFile;
                    dsd.IsHomogeneous = false;
                    dsd.LogFilePath = Path.Combine(this.outputDir, LOG);


                    PostProcessDSD(dsd);

                    return true;
                }
            }

            // Creates an entry collection (one per layout) for the DSD file
            private DsdEntryCollection CreateDsdEntryCollection(ArrayList layouts, Database db, string layerName)
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {

                    DsdEntryCollection entries = new DsdEntryCollection();

                    foreach (ObjectId layoutId in layouts)
                    {
                        Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

                        DsdEntry dsdEntry = new DsdEntry();
                        dsdEntry.DwgName = this.dwgFile;
                        dsdEntry.Layout = layout.LayoutName;
                        dsdEntry.Title = Path.GetFileNameWithoutExtension(this.dwgFile) + "-" + layerName + "-" + layout.LayoutName;
                        dsdEntry.Nps = layout.TabOrder.ToString();

                        entries.Add(dsdEntry);
                    }

                    tr.Commit();
                    return entries;
                }
            }

            // Writes the definitive DSD file from the templates and additional informations
            private void PostProcessDSD(DsdData dsd)
            {
                string str, newStr;
                string tmpFile = Path.Combine(this.outputDir, "temp.dsd");

                dsd.WriteDsd(tmpFile);


                using (StreamReader reader = new StreamReader(tmpFile, Encoding.Default))
                using (StreamWriter writer = new StreamWriter(this.dsdFile, false, Encoding.Default))
                {
                    while (!reader.EndOfStream)
                    {
                        str = reader.ReadLine();
                        if (str.Contains("Has3DDWF"))
                        {
                            newStr = "Has3DDWF=0";
                        }
                        else if (str.Contains("OriginalSheetPath"))
                        {
                            newStr = "OriginalSheetPath=" + this.dwgFile;
                        }
                        else if (str.Contains("Type"))
                        {
                            newStr = "Type=" + this.plotType;
                        }
                        else if (str.Contains("OUT"))
                        {
                            newStr = "OUT=" + this.outputDir;
                        }
                        else if (str.Contains("IncludeLayer"))
                        {
                            newStr = "IncludeLayer=FALSE";
                        }
                        else if (str.Contains("PromptForDwfName"))
                        {
                            newStr = "PromptForDwfName=FALSE";
                        }
                        else if (str.Contains("LogFilePath"))
                        {
                            newStr = "LogFilePath=" + Path.Combine(this.outputDir, LOG);
                        }
                        else
                        {
                            newStr = str;
                        }
                        writer.WriteLine(newStr);
                    }
                }
                File.Delete(tmpFile);
            }
        }



        // Class to plot one PDF file per sheet
        public class SingleSheetPdf : PlotToFileConfig
        {
            public SingleSheetPdf(string outputDir, ArrayList layouts)
                : base(outputDir, layouts, "5") { }
        }


    }
}
