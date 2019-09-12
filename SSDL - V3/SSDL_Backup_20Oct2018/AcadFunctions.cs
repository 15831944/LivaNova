using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;

using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Autodesk.AutoCAD.Colors;
using System.Windows.Forms;
using System.Xml;
using System.Collections;
using System.Text.RegularExpressions;

namespace SSDL
{
    public static class AcadFunctions
    {
        public static string partnoatt = "PARTNO";
        public static string qtyatt = "QTY";
        public static string itemnoatt = "ITEMNO";
        public static string lenatt = "LEN";

        public static void OpenDrawing(string strFileName)
        {

            DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

            if (File.Exists(strFileName))
            {
                Document doc = acDocMgr.Open(strFileName, false);
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = doc;
            }
            else
            {
                acDocMgr.MdiActiveDocument.Editor.WriteMessage("File " + strFileName +
                                                                " does not exist.");
            }
        }
        static public void ExtractObjectsFromFile(string dwgpath, out string partno, out string qty, out string itemno, out string len, out bool lencheck)
        {
            partno = "";
            qty = ""; itemno = ""; len = "";
            lencheck = false;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Create a database and try to load the file
            Database db = new Database(false, true);
            using (db)
            {
                try
                {
                    db.ReadDwgFile(dwgpath, System.IO.FileShare.Read, false, "");
                }
                catch (System.Exception)
                {
                    ed.WriteMessage("\nUnable to read Block file.");
                    return;
                }

                Transaction tr = db.TransactionManager.StartTransaction();
                using (tr)
                {
                    // Open the blocktable, get the modelspace
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);


                    // Iterate through it, dumping objects
                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);

                        if (ent is AttributeDefinition)
                        {
                            if (((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).Tag == partnoatt)
                            {
                                partno = ((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).TextString;
                            }
                            else if (((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).Tag == itemnoatt)
                            {
                                itemno = ((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).TextString;
                            }
                            else if (((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).Tag == qtyatt)
                            {
                                qty = ((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).TextString;
                            }
                            else if (((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).Tag == lenatt)
                            {
                                len = ((Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)ent).TextString;
                                lencheck = true;
                            }
                        }
                    }
                }

            }
        }

        public static void InsertBlk(string filepath, string partnoval, string qtyval, string itemnoval, string lenval)
        {

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (DocumentLock dl = doc.LockDocument())
            {
                for (int i = 0; i < Convert.ToInt32(qtyval); i++)
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        ObjectId btrId = bt.GetBlock(filepath);
                        if (btrId == ObjectId.Null)
                            return;
                        Dictionary<string, string> attValues = new Dictionary<string, string>();
                        attValues.Add(partnoatt, partnoval);
                        attValues.Add(qtyatt, "1");
                        // attValues.Add(itemnoatt, itemnoval);
                        attValues.Add(lenatt, lenval);
                        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        string blockName = System.IO.Path.GetFileNameWithoutExtension(filepath);

                        PromptPointOptions opPt = new PromptPointOptions("pick the insert point of the blk");
                        PromptPointResult resPt = default(PromptPointResult);
                        resPt = ed.GetPoint(opPt);
                        if (resPt.Status != PromptStatus.OK)
                        {
                            System.Windows.MessageBox.Show("fail to get the insert point");
                            return;
                        }

                        Point3d ptInsert = default(Point3d);
                        ptInsert = resPt.Value;

                        BlockReference br = curSpace.InsertBlockReference(blockName, ptInsert, attValues);
                        tr.Commit();
                    }
                }
            }
            return;

            //Database db = Application.DocumentManager.MdiActiveDocument.Database;
            //Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            //using (Transaction trans = db.TransactionManager.StartTransaction())
            //{
            //    try
            //    {
            //        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            //        if (bt != null)
            //        {
            //            BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            //            string blockname = System.IO.Path.GetFileNameWithoutExtension(filepath);

            //            ObjectId id = default(ObjectId);
            //            if (bt.Has(blockname))
            //            {
            //                BlockTableRecord btrSrc = trans.GetObject(bt[blockname], OpenMode.ForRead) as BlockTableRecord;
            //                id = btrSrc.Id;
            //            }
            //            else
            //            {
            //                Database dbDwg = new Database(false, true);
            //                dbDwg.ReadDwgFile(filepath, System.IO.FileShare.Read, true, "");
            //                id = db.Insert(blockname, dbDwg, true);
            //                dbDwg.Dispose();
            //            }
            //            PromptPointOptions opPt = new PromptPointOptions("pick the insert point of the blk");
            //            PromptPointResult resPt = default(PromptPointResult);
            //            resPt = ed.GetPoint(opPt);
            //            if (resPt.Status != PromptStatus.OK)
            //            {
            //                System.Windows.MessageBox.Show("fail to get the insert point");
            //                return;
            //            }
            //            Point3d ptInsert = default(Point3d);
            //            ptInsert = resPt.Value;
            //            BlockReference blkRef = new BlockReference(ptInsert , id);
            //            btr.AppendEntity(blkRef);
            //            trans.AddNewlyCreatedDBObject(blkRef, true);
            //            trans.Commit();


            //            UpdateAttributesInBlock(id, blockname, partnoatt, partnoval);
            //            UpdateAttributesInBlock(id, blockname, qtyatt, qtyval);
            //            UpdateAttributesInBlock(id, blockname, itemnoatt, itemnoval);
            //        }
            //    }
            //    catch
            //    {
            //        System.Windows.MessageBox.Show("fail to read dwg file and insert in current database ");
            //    }
            //    finally
            //    {
            //        trans.Dispose();
            //    }
            //}
        }

        private static int UpdateAttributesInBlock(ObjectId btrId, string blockName, string attbName, string attbValue)
        {
            // Will return the number of attributes modified

            int changedCount = 0;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            Transaction tr = doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                // Test each entity in the container...

                foreach (ObjectId entId in btr)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;

                    if (ent != null)
                    {
                        BlockReference br = ent as BlockReference;
                        if (br != null)
                        {
                            BlockTableRecord bd = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);

                            // ... to see whether it's a block with
                            // the name we're after

                            if (bd.Name.ToUpper() == blockName)
                            {

                                // Check each of the attributes...

                                foreach (ObjectId arId in br.AttributeCollection)
                                {
                                    DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                    AttributeReference ar =
                                      obj as AttributeReference;
                                    if (ar != null)
                                    {
                                        // ... to see whether it has
                                        // the tag we're after

                                        if (ar.Tag.ToUpper() == attbName)
                                        {
                                            // If so, update the value
                                            // and increment the counter

                                            ar.UpgradeOpen();
                                            ar.TextString = attbValue;
                                            ar.DowngradeOpen();
                                            changedCount++;
                                        }
                                    }
                                }
                            }

                            // Recurse for nested blocks
                            changedCount += UpdateAttributesInBlock(br.BlockTableRecord, blockName, attbName, attbValue);
                        }
                    }
                }
                tr.Commit();
            }
            return changedCount;
        }

        public static List<LayerTableRecord> AcadLayersList { get; set; }

        public static List<string> ActiveLayers { get; set; }

        public static string ActiveLayerName { get; set; }

        public static List<Layerlist> GetLayerlist()
        {
            List<Layerlist> result = new List<Layerlist>();

            AcadLayersList = new List<LayerTableRecord>();
            ActiveLayers = new List<string>();
            ActiveLayerName = string.Empty;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor aceditor = acDoc.Editor;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                // This example returns the layer table for the current database
                LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                // Step through the Layer table and print each layer name
                foreach (ObjectId acObjId in acLyrTbl)
                {
                    Layerlist ls = new SSDL.Layerlist();
                    LayerTableRecord acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;

                    if (acObjId == acCurDb.Clayer)
                    {
                        ActiveLayerName = acLyrTblRec.Name;
                    }

                    string layerName = acLyrTblRec.Name;
                    ls.LayerName = layerName;

                    TypedValue[] tvs = new TypedValue[1] { new TypedValue((int)DxfCode.LayerName, layerName) };
                    SelectionFilter sf = new SelectionFilter(tvs);
                    PromptSelectionResult psr = aceditor.SelectAll(sf);
                    ObjectIdCollection entcollection = new ObjectIdCollection();

                    if (psr.Status == PromptStatus.OK)
                    {
                        entcollection = new ObjectIdCollection(psr.Value.GetObjectIds());
                    }

                    int ViewPortCount = 0;
                    //using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                    {
                        for (int i = 0; i < entcollection.Count; i++)
                        {
                            Entity enty = acTrans.GetObject(entcollection[i], OpenMode.ForRead) as Entity;
                            if (enty is Viewport)
                            {
                                //entcollection.Remove(entcollection[i]);
                                ViewPortCount++;
                            }
                        }
                        // trans.Commit();
                    }
                    if (entcollection.Count - ViewPortCount <= 0)
                    {
                        ls.EntityAvail = false;
                    }
                    else
                    {
                        ls.EntityAvail = true;
                    }

                    ls.Include = false;
                    result.Add(ls);
                    AcadLayersList.Add(acLyrTblRec);
                    if (!acLyrTblRec.IsOff)
                    {
                        ActiveLayers.Add(acLyrTblRec.Name);
                    }

                    result = result.OrderBy(x => PadNumbers(x.LayerName)).ToList();

                }
            }

            // Dispose of the transaction

            return result;
        }

        public static string PadNumbers(string input)
        {
            return Regex.Replace(input, "[0-9]+", match => match.Value.PadLeft(10, '0'));
        }
    }

    public class AlphanumComparatorFast : IComparer
    {
        public int Compare(object x, object y)
        {
            string s1 = x as string;
            if (s1 == null)
            {
                return 0;
            }
            string s2 = y as string;
            if (s2 == null)
            {
                return 0;
            }

            int len1 = s1.Length;
            int len2 = s2.Length;
            int marker1 = 0;
            int marker2 = 0;

            // Walk through two the strings with two markers.
            while (marker1 < len1 && marker2 < len2)
            {
                char ch1 = s1[marker1];
                char ch2 = s2[marker2];

                // Some buffers we can build up characters in for each chunk.
                char[] space1 = new char[len1];
                int loc1 = 0;
                char[] space2 = new char[len2];
                int loc2 = 0;

                // Walk through all following characters that are digits or
                // characters in BOTH strings starting at the appropriate marker.
                // Collect char arrays.
                do
                {
                    space1[loc1++] = ch1;
                    marker1++;

                    if (marker1 < len1)
                    {
                        ch1 = s1[marker1];
                    }
                    else
                    {
                        break;
                    }
                } while (char.IsDigit(ch1) == char.IsDigit(space1[0]));

                do
                {
                    space2[loc2++] = ch2;
                    marker2++;

                    if (marker2 < len2)
                    {
                        ch2 = s2[marker2];
                    }
                    else
                    {
                        break;
                    }
                } while (char.IsDigit(ch2) == char.IsDigit(space2[0]));

                // If we have collected numbers, compare them numerically.
                // Otherwise, if we have strings, compare them alphabetically.
                string str1 = new string(space1);
                string str2 = new string(space2);

                int result;

                if (char.IsDigit(space1[0]) && char.IsDigit(space2[0]))
                {
                    int thisNumericChunk = int.Parse(str1);
                    int thatNumericChunk = int.Parse(str2);
                    result = thisNumericChunk.CompareTo(thatNumericChunk);
                }
                else
                {
                    result = str1.CompareTo(str2);
                }

                if (result != 0)
                {
                    return result;
                }
            }
            return len1 - len2;
        }
    }

    public class AcadCommands
    {
        [CommandMethod("opendrawing")]
        public void opendrawing()
        {
            MainWindow od = new MainWindow();
            od.ShowDialog();
        }

        [CommandMethod("OpenSymbolLibrary")]
        public void OpenSymbolLibrary()
        {
            BlockView od = new BlockView();
            od.ShowDialog();
        }

        [CommandMethod("GenerateBOMLayer")]
        public void GenerateBOMLayer()
        {
            BOM od = new BOM();
            od.Generate(AutocadLayerwise: true);
        }

        [CommandMethod("GenerateBOM")]
        public void GenerateBOM()
        {
            BOM od = new BOM();
            od.Generate(placeInAutoCAD: true);
        }

        [CommandMethod("PrintBOM")]
        public void PrintBOM()
        {
            View.ProgressBar1 _pBar = new View.ProgressBar1();
            _pBar.Show();
            
            try
            {
                BOM od = new BOM();
                od.Generate(ExcelOutput: true);
            }
            catch { }
            _pBar.Close();
        }

        [CommandMethod("SendBOM")]
        public void SendBOM()
        {
            DialogResult dr = MessageBox.Show("Please do make sure description is selectable. Click Yes to proceed or No to cancel", "SSDL", MessageBoxButtons.YesNo);
            if (dr == DialogResult.No)
            {
                return;
            }
            BOM od = new BOM();

            od.SendXml();
        }

        [CommandMethod("PrintTimeRollUp")]
        public void PrintTimeRollUp()
        {
            BOM od = new BOM();

            System.Data.DataTable dtnew = od.Get_Drawingdata();

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
            //29-06-2018
            //DialogResult dr = MessageBox.Show("Please do make sure description is selectable. Click Yes to proceed or No to cancel", "SSDL", MessageBoxButtons.YesNo);
            //if (dr == DialogResult.No)
            //{
            //    return;
            //}

            string sumresult = "", drgnumber = "", SurfaceAreaSum = "";
            string projectDesc = od.getProjectDesc();
            bool result = od.timerollupfunctionality(out sumresult, out drgnumber, out SurfaceAreaSum, true, readDesc: false, projectDesc: projectDesc, divideBy60: false, updateDWGAttr: false); //02-05-2018

            //bool result = od.timerollupfunctionality(out sumresult, out drgnumber, out SurfaceAreaSum, readDesc: false, projectDesc: projectDesc);//02-05-2018

            //// xmlbuilder(dtnewExcel, sumresult, drgnumber, description, drgnumber, drgnumber);
            //System.Data.DataTable dtnew = od.Get_Drawingdata();
            //System.Data.DataTable dtnewExcel = dtnew.DefaultView.ToTable(false, "Itemno", "Quantity", "Partno", "Temp", "Description");
            //drgnumber = drgnumber.Replace("P_", "").Trim();
            //od.xmlbuilder(dtnewExcel, sumresult, drgnumber, description, drgnumber, drgnumber);
            // if (result)
            od.SendXml(false, projectDesc, sumresult, drgnumber, SurfaceAreaSum);

            //Launch a shell script
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string canIMPORT = getpath.IniReadValue("FilePath", "IMPORT");
            string xmloutputLocation = getpath.IniReadValue("FilePath", "XML_OUTPUT_LOCATION"); ;
            if (canIMPORT.Equals("ON", StringComparison.InvariantCultureIgnoreCase))//15-05-2018, added to avoid message when Import is turned off
            {
                if (MessageBox.Show("Do you wish to Import The Time Rollup into Exact?", "SSDL", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        string EXACT_PATH = getpath.IniReadValue("ExactConfiguration", "EXACT_PATH");// C:\Program Files (x86)\Exact Software\bin\;
                        string EXACT_EXE_NAME = getpath.IniReadValue("ExactConfiguration", "EXACT_EXE_NAME");// AsImport.exe;
                        string EXACT_ARG1 = getpath.IniReadValue("ExactConfiguration", "EXACT_ARG1");//-rUS - AUS - PRDExact;
                        string EXACT_SERVER_ENVIRONMENT = getpath.IniReadValue("ExactConfiguration", "EXACT_SERVER_ENVIRONMENT");//-D999;
                        string EXACT_ARG2 = getpath.IniReadValue("ExactConfiguration", "EXACT_ARG2");//-u - ~I;
                        string EXACT_ARG3 = getpath.IniReadValue("ExactConfiguration", "EXACT_ARG3");//-TITEMS;
                        string EXACT_ARG4 = getpath.IniReadValue("ExactConfiguration", "EXACT_ARG4");//-OAuto;

                        string shellCommand = EXACT_PATH + EXACT_EXE_NAME + " " + EXACT_ARG1 + " " + EXACT_SERVER_ENVIRONMENT + " " + EXACT_ARG2 + " -URL" + xmloutputLocation + " " + EXACT_ARG3 + " " + EXACT_ARG4;

                        ExecuteCommandSync(shellCommand);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        /// <span class="code-SummaryComment"><summary></span>
        /// Executes a shell command synchronously.
        /// <span class="code-SummaryComment"></summary></span>
        /// <span class="code-SummaryComment"><param name="command">string command</param></span>
        /// <span class="code-SummaryComment"><returns>string, as output of the command.</returns></span>
        public void ExecuteCommandSync(object command)
        {
            try
            {
                // create the ProcessStartInfo using "cmd" as the program to be run,
                // and "/c " as the parameters.
                // Incidentally, /c tells cmd that we want it to execute the command that follows,
                // and then exit.
                System.Diagnostics.ProcessStartInfo procStartInfo =
                    new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);

                // The following commands are needed to redirect the standard output.
                // This means that it will be redirected to the Process.StandardOutput StreamReader.
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = true;
                // Now we create a process, assign its ProcessStartInfo and start it
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
                // Get the output into a string
                string result = proc.StandardOutput.ReadToEnd();
                // Display the command output.
                Console.WriteLine(result);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Shell command execution failed - " + ex.ToString(), "SSDL", MessageBoxButtons.OK);
            }
        }

        [CommandMethod("PDFLayerCreation", CommandFlags.Session)]
        public void PDFLayerCreation()
        {
            PDFLayerview pdfview = new PDFLayerview(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
            pdfview.ShowDialog();
        }
    }

    public static class DatabaseClass
    {

        public static System.Data.DataTable GetProductsRM()
        {
            // string query = "SELECT  *, RIGHT('000000000'+CAST(RM_No AS VARCHAR(9)),9) from Products_RM";
            //string query = "SELECT *, Format(RM_No, '000000000) as 'RM_No1' from Products_RM";
            string sFormat = "000000000";
            //, [RM_Desc], [Vendor_Num], [UoM], [Std_Cost], [fldTimeRollup], [Surface_Area] 
            string query = "SELECT Format(RM_No, '000000000') as [RM_No], [RM_Desc], [Vendor_Num], [UoM], [Std_Cost], [fldTimeRollup], [Surface_Area]  from Products_RM";
            return getFromDataBase(query);// "SELECT * from Products_RM");
        }

        public static System.Data.DataTable GetTimeRollupCategories()
        {
            return getFromDataBase("SELECT * from tblTime_Rollup_Categories order by fldCategory");
        }

        public static System.Data.DataTable getFromDataBase(string query)
        {
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string mdb = getpath.IniReadValue("FilePath", "SERVER_BOM_DATABASE_PATH");
            //= Properties.Settings.Default.SERVER_BOM_DATABASE_PATH;
            string constring = getpath.IniReadValue("FilePath", "ConnectionString");

            System.Data.DataTable dtProducts_RM = new System.Data.DataTable();
            using (OleDbConnection connection = new OleDbConnection(constring))
            {
                if (connection.State == System.Data.ConnectionState.Open)
                    connection.Close();
                connection.Open();
                //Adjust your query

                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection); //This assigns the Select statement and connection of the data adapter 
                //MessageBox.Show(query + " - 1" + connection.ToString());
                OleDbCommandBuilder oleDbCommandBuilder = new OleDbCommandBuilder(adapter); //This builds the update and Delete queries for the table in the above SQL. this only works if the select is a single table. 
                                                                                            // MessageBox.Show(query + " - 2" + connection.ToString());
                adapter.Fill(dtProducts_RM);
                // MessageBox.Show(query + " - 3" + connection.ToString());

            }
            return dtProducts_RM;
        }



    }


    public class BOM
    {
        public List<KeyValuePair<string, string>> sheetvalue = new List<KeyValuePair<string, string>>();


        public void updateDwgData(Dictionary<int, string> dicItemNo)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Build a filter list so that only
                        // block references are selected
                        TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
                        SelectionFilter filter = new SelectionFilter(filList);
                        PromptSelectionOptions opts = new PromptSelectionOptions();
                        PromptSelectionResult res = ed.SelectAll(filter);

                        // Do nothing if selection is unsuccessful
                        if (res.Status != PromptStatus.OK)
                            return;

                        SelectionSet selSet = res.Value;
                        ObjectId[] idArray = selSet.GetObjectIds();
                        foreach (ObjectId blkId in idArray)
                        {
                            BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);
                            AttributeCollection attCol = blkRef.AttributeCollection;
                            string partNo = string.Empty;
                            string len = string.Empty;
                            bool lenavail = false;
                            foreach (ObjectId attId in attCol)
                            {
                                using (AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead))
                                {
                                    if (attRef.Tag.Equals(AcadFunctions.partnoatt, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        partNo = attRef.TextString;
                                        //break;
                                    }
                                    else if (attRef.Tag.Equals(AcadFunctions.lenatt, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        lenavail = true;
                                        len = attRef.TextString;
                                    }
                                }
                            }
                            if (lenavail)
                            {
                                if (lenavail)
                                {
                                    partNo = partNo.Substring(0, 6) + len.PadLeft(3, '0');
                                }
                            }

                            if (!string.IsNullOrEmpty(partNo))
                            {
                                foreach (ObjectId attId in attCol)
                                {
                                    try
                                    {
                                        using (AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite))
                                        {
                                            if (attRef.Tag.Equals(AcadFunctions.itemnoatt, StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.TextString = dicItemNo.FirstOrDefault(a => a.Value.ToString().Equals(partNo, StringComparison.InvariantCultureIgnoreCase)).Key.ToString();
                                                break;
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    catch { }
                    tr.Commit();
                }
            }
            catch { }
        }

        public System.Data.DataTable Get_Drawingdata(bool includeLayer = false, bool includeLenAttr = true, bool concatLengthWithPartNo = true, bool sendXML = false, bool updateDwgAttr = true)
        {
            Cursor.Current = Cursors.WaitCursor;
            System.Data.DataTable dt = DatabaseClass.GetProductsRM();

            System.Data.DataTable dtnewExcel = new System.Data.DataTable();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;


            System.Data.DataTable dtnew = new System.Data.DataTable();
            dtnew.Columns.Add("Itemno");
            dtnew.Columns.Add("Partno");
            dtnew.Columns.Add("Quantity", typeof(double));// System.Type.GetType("System.Int32"));
            dtnew.Columns.Add("Description");
            dtnew.Columns.Add("Timerollup", System.Type.GetType("System.Int32"));
            dtnew.Columns.Add("SurfaceArea", typeof(double));
            if (includeLayer)
                dtnew.Columns.Add("Layer");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the blocktable, get the modelspace
                // BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                // BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                string partno = "", itemno = "", qty = "", desc = "", timerollup = "", len = "", SurfaceArea = "";
                bool lenavail = false;
                System.Data.DataRow dr = null;

                try
                {
                    // Build a filter list so that only
                    // block references are selected
                    TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
                    SelectionFilter filter = new SelectionFilter(filList);
                    PromptSelectionOptions opts = new PromptSelectionOptions();
                    PromptSelectionResult res = ed.SelectAll(filter);

                    // Do nothing if selection is unsuccessful
                    if (res.Status != PromptStatus.OK)
                        return null;

                    SelectionSet selSet = res.Value;
                    ObjectId[] idArray = selSet.GetObjectIds();
                    foreach (ObjectId blkId in idArray)
                    {
                        BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);
                        // btr.Dispose();
                        lenavail = false;
                        if (blkRef == null) continue;
                        AttributeCollection attCol = blkRef.AttributeCollection;

                        if (blkRef.Name.Equals("SheetFormat", StringComparison.InvariantCultureIgnoreCase) || blkRef.Name.StartsWith("Title", StringComparison.InvariantCultureIgnoreCase))
                        {
                            foreach (ObjectId attId in attCol)
                            {
                                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                                if (attRef.Tag.Equals("ECO", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    sheetvalue.Add(new KeyValuePair<string, string>("ECO", attRef.TextString));
                                }
                                else if (attRef.Tag.Equals("SHT", StringComparison.InvariantCultureIgnoreCase) || attRef.Tag.Equals("PAGES", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    sheetvalue.Add(new KeyValuePair<string, string>("SHT", attRef.TextString));
                                }
                                else if (attRef.Tag.Equals("DRG", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    sheetvalue.Add(new KeyValuePair<string, string>("DRG", attRef.TextString));
                                }
                                else if (attRef.Tag.Equals("REV", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    sheetvalue.Add(new KeyValuePair<string, string>("REV", attRef.TextString));
                                }
                            }
                        }
                        else
                        {
                            foreach (ObjectId attId in attCol)
                            {
                                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);

                                if (attRef.Tag.Equals(AcadFunctions.partnoatt, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    partno = attRef.TextString;
                                }
                                else if (attRef.Tag.Equals(AcadFunctions.itemnoatt, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    // itemno = attRef.TextString;
                                }
                                else if (attRef.Tag.Equals(AcadFunctions.qtyatt, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    if (string.IsNullOrEmpty(attRef.TextString) || attRef.TextString.Equals("0", StringComparison.InvariantCultureIgnoreCase))
                                        qty = "1";// "0"; - Changed to count the block instead of considering qty attribute - 01 March 2018
                                    else
                                        qty = attRef.TextString;
                                }
                                else if (attRef.Tag.Equals(AcadFunctions.lenatt, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    lenavail = true;
                                    len = attRef.TextString;
                                }
                            }
                            try
                            {
                                //dr = dt.Select("RM_No =" + partno)[0];
                                //if (dr == null || lenavail)
                                //    dr = dt.Select("RM_No =" + blkRef.Name)[0];

                                DataRow[] drr = dt.Select("RM_No ='" + partno + "' or RM_No = '" + blkRef.Name + "'");
                                if (drr.Count() > 0)
                                {
                                    dr = drr[0];
                                    desc = dr["RM_Desc"].ToString();
                                }

                                if (dr["fldTimeRollup"] != null)
                                {
                                    if (!string.IsNullOrEmpty(dr["fldTimeRollup"].ToString()))
                                        timerollup = dr["fldTimeRollup"].ToString();
                                    else
                                        timerollup = "0";
                                }
                                else
                                    timerollup = "0";

                                if (dr["Surface_Area"] != null)
                                {
                                    if (!string.IsNullOrEmpty(dr["Surface_Area"].ToString()))
                                        SurfaceArea = dr["Surface_Area"].ToString();
                                    else
                                        SurfaceArea = "0";
                                }


                                if (lenavail && includeLenAttr)
                                {
                                    qty = len;
                                    //timerollup = len;
                                }
                                if (lenavail && concatLengthWithPartNo)
                                {
                                    partno = partno.Substring(0, 6) + len.PadLeft(3, '0');
                                }
                                else if (lenavail && !concatLengthWithPartNo && qty.Length > 0)
                                    qty = (Convert.ToDouble(qty) / 100).ToString();

                                //Commented by Sundari in advice of rajasekar for the Bushing error given by Darren on 16-08-2018
                                if (sendXML && lenavail)//15-05-2018, added to use block name in case of Tube while writing to xml
                                {
                                    try
                                    {
                                        long _iPartNo = 0;
                                        bool isLong = long.TryParse(blkRef.Name.ToString(), out _iPartNo);
                                        if (isLong && _iPartNo > 0)
                                            partno = blkRef.Name;
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                desc = "";
                                timerollup = "0";
                                SurfaceArea = "0";
                                if (string.IsNullOrEmpty(qty) || qty == "0")
                                    qty = "1";
                            }
                            try
                            {
                                if (!string.IsNullOrEmpty(partno))//added on 14-8-2018, to ignore blocks with out part no
                                {
                                    if (includeLayer)
                                        dtnew.Rows.Add(itemno, partno, qty, desc, timerollup, Math.Round(Convert.ToDouble(SurfaceArea) * Convert.ToDouble(qty), 3), blkRef.Layer);
                                    else
                                        dtnew.Rows.Add(itemno, partno, qty, desc, timerollup, Math.Round(Convert.ToDouble(SurfaceArea) * Convert.ToDouble(qty), 3));
                                }
                            }
                            catch { }
                        }

                        partno = ""; itemno = ""; qty = ""; desc = ""; timerollup = ""; len = ""; SurfaceArea = "";
                    }
                    tr.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    //ed.WriteMessage(("Exception: " + ex.Message));
                }

                //Sort by part no and update Item No
                System.Data.DataTable distinctValues = dtnew.DefaultView.ToTable(true, "Partno");
                Dictionary<int, string> dicItemNo = new Dictionary<int, string>();
                int val = 1;
                foreach (DataRow drVal in distinctValues.Rows)
                {
                    dicItemNo.Add(val, drVal["Partno"].ToString());
                    val++;
                }
                if (updateDwgAttr)
                    updateDwgData(dicItemNo);
                foreach (DataRow drVal in dtnew.Rows)
                {
                    try
                    {
                        drVal["ItemNo"] = dicItemNo.FirstOrDefault(a => a.Value.ToString().Equals(drVal["Partno"].ToString(), StringComparison.InvariantCultureIgnoreCase)).Key;
                    }
                    catch { }
                }


                if (dtnew.Rows.Count > 0)
                {
                    if (includeLayer)
                    {
                        var result = dtnew.AsEnumerable()
                         .GroupBy(x => new { layer = x.Field<string>("Layer"), item = x.Field<string>("Itemno"), part = x.Field<string>("Partno"), desc = x.Field<string>("Description") })
                         .Select(x => new
                         {
                             x.Key.item,
                             x.Key.part,
                             TotalSum = x.Sum(z => z.Field<double>("Quantity")),
                             x.Key.desc,
                             Timesum = x.Sum(z => z.Field<int>("Timerollup")),
                             SurfaceArea = x.Sum(z => z.Field<double>("SurfaceArea")),
                             x.Key.layer
                         }).ToList();



                        dtnewExcel.Columns.Add("Itemno", typeof(Int32));
                        dtnewExcel.Columns.Add("Quantity", typeof(double));// System.Type.GetType("System.Int32"));
                        dtnewExcel.Columns.Add("Partno");
                        dtnewExcel.Columns.Add("Temp");
                        dtnewExcel.Columns.Add("Description");
                        dtnewExcel.Columns.Add("Timerollup");
                        dtnewExcel.Columns.Add("SurfaceArea", typeof(double));
                        dtnewExcel.Columns.Add("Layer");

                        foreach (var item in result)
                        {
                            try
                            {
                                if (item.item == "" || item.item == null)
                                    dtnewExcel.Rows.Add("0", item.TotalSum, item.part, "", item.desc, item.Timesum, item.SurfaceArea, item.layer);
                                else
                                    dtnewExcel.Rows.Add(item.item, item.TotalSum, item.part, "", item.desc, item.Timesum, item.SurfaceArea, item.layer);
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }

                        DataView dv = dtnewExcel.DefaultView;
                        dv.Sort = "Itemno";
                        dtnewExcel = dv.ToTable();
                    }
                    else
                    {
                        var result = dtnew.AsEnumerable()
                        .GroupBy(x => new { item = x.Field<string>("Itemno"), part = x.Field<string>("Partno"), desc = x.Field<string>("Description") })
                        .Select(x => new
                        {
                            x.Key.item,
                            x.Key.part,
                            TotalSum = x.Sum(z => z.Field<double>("Quantity")),
                            x.Key.desc,
                            Timesum = x.Sum(z => z.Field<int>("Timerollup")),
                            SurfaceArea = x.Sum(z => z.Field<double>("SurfaceArea"))
                        }).ToList();


                        dtnewExcel.Columns.Add("Itemno", typeof(Int32));
                        dtnewExcel.Columns.Add("Quantity", typeof(double));// System.Type.GetType("System.Int32"));
                        dtnewExcel.Columns.Add("Partno");
                        dtnewExcel.Columns.Add("Temp");
                        dtnewExcel.Columns.Add("Description");
                        dtnewExcel.Columns.Add("Timerollup");
                        dtnewExcel.Columns.Add("SurfaceArea", typeof(double));

                        foreach (var item in result)
                        {
                            try
                            {
                                if (item.item == "" || item.item == null)
                                    dtnewExcel.Rows.Add("0", item.TotalSum, item.part, "", item.desc, item.Timesum, item.SurfaceArea);
                                else
                                    dtnewExcel.Rows.Add(item.item, item.TotalSum, item.part, "", item.desc, item.Timesum, item.SurfaceArea);
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }

                        DataView dv = dtnewExcel.DefaultView;
                        dv.Sort = "Itemno, Partno";//Added Partno on 30-04-2018
                        dtnewExcel = dv.ToTable();
                    }
                }
            }
            Cursor.Current = Cursors.Default;
            return dtnewExcel;
        }

        private Point3d Get_Max_Min_Point(string Layer_name, string Bom_Location)
        {

            Point3d insertpoint = new Point3d();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            System.Data.DataTable dtpoint = new System.Data.DataTable();
            dtpoint.Columns.Add("MinX", typeof(double));
            dtpoint.Columns.Add("MinY", typeof(double));
            dtpoint.Columns.Add("MaxX", typeof(double));// System.Type.GetType("System.Int32"));
            dtpoint.Columns.Add("MaxY", typeof(double));

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Build a filter list so that only
                    // block references are selected
                    TypedValue[] filList = new TypedValue[1] { new TypedValue((int)DxfCode.Start, "INSERT") };
                    SelectionFilter filter = new SelectionFilter(filList);
                    PromptSelectionOptions opts = new PromptSelectionOptions();
                    PromptSelectionResult res = ed.SelectAll(filter);

                    // Do nothing if selection is unsuccessful
                    if (res.Status != PromptStatus.OK)
                        return insertpoint;

                    SelectionSet selSet = res.Value;
                    ObjectId[] idArray = selSet.GetObjectIds();
                    foreach (ObjectId blkId in idArray)
                    {
                        try
                        {
                            Entity _ent = (Entity)tr.GetObject(blkId, OpenMode.ForRead);
                            if (Layer_name.Equals(_ent.Layer, StringComparison.InvariantCultureIgnoreCase))
                                dtpoint.Rows.Add(_ent.GeometricExtents.MinPoint.X, _ent.GeometricExtents.MinPoint.Y, _ent.GeometricExtents.MaxPoint.X, _ent.GeometricExtents.MaxPoint.Y);

                        }
                        catch { }
                        //BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);
                        //// btr.Dispose();
                        //if (Layer_name.Equals(blkRef.Layer, StringComparison.InvariantCultureIgnoreCase))
                        //{
                        //    dtpoint.Rows.Add(blkRef.GeometricExtents.MinPoint.X, blkRef.GeometricExtents.MinPoint.Y, blkRef.GeometricExtents.MaxPoint.X, blkRef.GeometricExtents.MaxPoint.Y);
                        //}
                    }
                    tr.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    //ed.WriteMessage(("Exception: " + ex.Message));
                }

                if (Bom_Location.Equals("TOPLEFT", StringComparison.InvariantCultureIgnoreCase))
                {
                    insertpoint = new Point3d(Convert.ToDouble(dtpoint.Compute("min([MinX])", string.Empty)), Convert.ToDouble(dtpoint.Compute("max([MaxY])", string.Empty)), 0);
                }
                else if (Bom_Location.Equals("TOPRIGHT", StringComparison.InvariantCultureIgnoreCase))
                {
                    insertpoint = new Point3d(Convert.ToDouble(dtpoint.Compute("max([MaxX])", string.Empty)), Convert.ToDouble(dtpoint.Compute("max([MaxY])", string.Empty)), 0);
                }
                else if (Bom_Location.Equals("BOTTOMLEFT", StringComparison.InvariantCultureIgnoreCase))
                {
                    insertpoint = new Point3d(Convert.ToDouble(dtpoint.Compute("min([MinX])", string.Empty)), Convert.ToDouble(dtpoint.Compute("min([MinY])", string.Empty)), 0);
                }
                else if (Bom_Location.Equals("BOTTOMRIGHT", StringComparison.InvariantCultureIgnoreCase))
                {
                    insertpoint = new Point3d(Convert.ToDouble(dtpoint.Compute("max([MaxX])", string.Empty)), Convert.ToDouble(dtpoint.Compute("min([MinY])", string.Empty)), 0);
                }

            }
            return insertpoint;
        }

        public void SendXml(bool readDesc = true, string description = "", string sumresult = "", string drgnumber = "", string SurfaceAreaSum = "")
        {
            try
            {
                System.Data.DataTable dtnew = Get_Drawingdata(concatLengthWithPartNo: false, sendXML: true, updateDwgAttr: false);
                //Changed by Chockalingam on 17-08-2018
                //System.Data.DataTable dtnew = Get_Drawingdata(concatLengthWithPartNo: true, sendXML: true, updateDwgAttr: false);
                if (dtnew == null)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                if (dtnew.Rows.Count <= 0)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }

                System.Data.DataTable dtnewExcel = dtnew.DefaultView.ToTable(false, "Itemno", "Quantity", "Partno", "Temp", "Description", "SurfaceArea");

                if (readDesc)
                {
                    bool result = timerollupfunctionality(out sumresult, out drgnumber, out SurfaceAreaSum, false, divideBy60: false);
                    if (!result) return;
                }
                drgnumber = drgnumber.Replace("P_", "").Trim();

                if (readDesc)
                    description = getProjectDesc();
                xmlbuilder(dtnewExcel, sumresult, drgnumber, description, drgnumber, drgnumber, SurfaceAreaSum);

            }
            catch (System.Exception ex)
            {
                // MessageBox.Show("XML Generation failed - " + ex.ToString());
            }
        }

        public static void xmlbuilder(System.Data.DataTable bom, string totaltimerollup, string itemcode = "", string description = "", string Bomcode = "", string drawingno = "", string SurfaceArea = "")
        {
            IniFile getpath1 = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string maxChar = getpath1.IniReadValue("FilePath", "MAX_BOM_CHAR");

            int maxCharAllowed = (string.IsNullOrEmpty(maxChar)) ? 0 : Convert.ToInt16(maxChar);

            StringBuilder stbuild = new StringBuilder();
            if (description.Length > maxCharAllowed)//60)// 30) // 10-05-2018//29-06-2018
                description = description.Substring(0, maxCharAllowed);// 60);// 30);// 10-05-2018//29-06-2018
            //stbuild.Append(@"<? xml version = ""1.0"" ?>").AppendLine();
            stbuild.Append(@"<eExact xsi:noNamespaceSchemaLocation=""eExact - Schema.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">").AppendLine();

            stbuild.Append("<Items>").AppendLine();
            //stbuild.Append("<Item code = \"" + itemcode + "\" type = \"S\" searchcode = \"6\">").AppendLine();
            stbuild.Append("<Item code = \"" + drawingno + "\">").AppendLine();
            stbuild.Append(@"<Condition>A</Condition>").AppendLine();

            stbuild.Append(@"<BOMs>").AppendLine();
            stbuild.Append("<BOM code=\"" + Bomcode + "\" versionnumber=\"1\">").AppendLine();
            stbuild.Append("<Description>" + description.Replace("&", "&amp;").Replace("'", "&apos;").Replace("\"", "&quot;").Replace(">", "&gt;").Replace("<", "&lt;") + "</Description>").AppendLine();
            stbuild.Append(@"<Linetype/>").AppendLine();  //stbuild.Append(@"<Linetype>1</Linetype>").AppendLine();
            stbuild.Append(@"<MainVersion>1</MainVersion>").AppendLine();
            stbuild.Append(@"<Condition>N</Condition>").AppendLine();
            stbuild.Append(@"<EffectiveDate>" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "</EffectiveDate>").AppendLine();//10-05-2018
            //stbuild.Append(@"<EffectiveDate>2006-01-01</EffectiveDate>").AppendLine();
            stbuild.Append(@"<DrawingNo>" + drawingno + "</DrawingNo>").AppendLine();
            stbuild.Append(@"<Quantity>1</Quantity>").AppendLine();
            //stbuild.Append(@"<CostCenter code=""MD1""/>").AppendLine();//30-04-2018
            //stbuild.Append(@"<CostCenter code=""001CC001""><Description>Default cost center</Description></CostCenter>").AppendLine();//30-04-2018
            stbuild.Append(@"<Costcenter code=""MD1""></Costcenter>").AppendLine();//05-05-2018
            //stbuild.Append(@"<Quantity> 1 </Quantity>").AppendLine();

            int ii = 1;
            foreach (DataRow dr in bom.Rows)
            {
                stbuild.Append("<BOMLine type=\"I\" sequencenumber=\"" + ii + "\">").AppendLine();
                stbuild.Append("<Item code=\"" + dr[2].ToString() + "\"/>").AppendLine();
                stbuild.Append(@"<Condition>N</Condition>").AppendLine();
                stbuild.Append(@"<Warehouse code=""KITR""/>").AppendLine();
                stbuild.Append(@"<BackFlush>0</BackFlush>").AppendLine();
                stbuild.Append(@"<EffectiveDate>" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "</EffectiveDate>").AppendLine();//10-05-2018
                //stbuild.Append(@"<EffectiveDate>2000-01-01</EffectiveDate>").AppendLine();
                // stbuild.Append(@"<ExpiryDate>2000-01-01</ExpiryDate>").AppendLine();//10-05-2018
                stbuild.Append(@"<Costcenter code=""MD1""/>").AppendLine();//30-04-2018

                stbuild.Append(@"<Quantity>" + Convert.ToDouble(dr[1].ToString()).ToString("F8") + "</Quantity>").AppendLine();
                stbuild.Append(@"</BOMLine>").AppendLine();

                ii++;
            }

            //Time Roll up in Hours
            stbuild.Append("<BOMLine type=\"L\" sequencenumber=\"" + ii + "\">").AppendLine();
            stbuild.Append("<Item code=\"PRODUCTION HOURS\"/>").AppendLine();
            stbuild.Append(@"<Condition>N</Condition>").AppendLine();
            stbuild.Append(@"<Warehouse code=""KITR""/>").AppendLine();
            stbuild.Append(@"<BackFlush>1</BackFlush>").AppendLine();
            stbuild.Append(@"<EffectiveDate>" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "</EffectiveDate>").AppendLine();//10-05-2018
            //stbuild.Append(@"<EffectiveDate>2000-01-01</EffectiveDate>").AppendLine();
            //stbuild.Append(@"<ExpiryDate>2000-01-01</ExpiryDate>").AppendLine();//10-05-2018
            //30-04-2018
            stbuild.Append(@"<Costcenter code=""MD1""/>").AppendLine();
            //stbuild.Append(@"<CostCenter code=""001CC001""><Description>Default cost center</Description></CostCenter>").AppendLine();//30-04-2018
            // stbuild.Append(@"<Quantity>" + Math.Round(Convert.ToDouble(totaltimerollup) / 60, 2).ToString() + "</Quantity>").AppendLine();//15-05-2018
            stbuild.Append(@"<Quantity>" + Math.Round(Math.Round(Convert.ToDouble(totaltimerollup) / 60, 2) / 60, 2).ToString() + "</Quantity>").AppendLine(); // 22-05-2018, Secs => Mins, Round to nearest value and divide by 60
            stbuild.Append(@"</BOMLine>").AppendLine();
            ii++;
            //Time Roll up in Minutes
            stbuild.Append("<BOMLine type=\"M\" sequencenumber=\"" + ii + "\">").AppendLine();
            stbuild.Append("<Item code=\"OVERHEAD\"/>").AppendLine();
            stbuild.Append(@"<Condition>N</Condition>").AppendLine();
            stbuild.Append(@"<Warehouse code=""KITR""/>").AppendLine();
            stbuild.Append(@"<BackFlush>1</BackFlush>").AppendLine();
            stbuild.Append(@"<EffectiveDate>" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "</EffectiveDate>").AppendLine();//10-05-2018
            //stbuild.Append(@"<EffectiveDate>2000-01-01</EffectiveDate>").AppendLine();
            //stbuild.Append(@"<ExpiryDate>2000-01-01</ExpiryDate>").AppendLine();//10-05-2018
            stbuild.Append(@"<Costcenter code=""MD1""/>").AppendLine();//30-04-2018
                                                                       //stbuild.Append(@"<CostCenter code=""001CC001""><Description>Default cost center</Description></CostCenter>").AppendLine();//30-04-2018

            //stbuild.Append(@"<Quantity>" + totaltimerollup + "</Quantity>").AppendLine();//17-05-2018
            if (totaltimerollup == "0")//22-5-2018
            {
                stbuild.Append(@"<Quantity>0</Quantity>").AppendLine();
            }
            else
            {
                stbuild.Append(@"<Quantity>" + Math.Round(Convert.ToDouble(totaltimerollup) / 60, 2).ToString() + "</Quantity>").AppendLine();//15-05-2018
            }
            stbuild.Append(@"</BOMLine>").AppendLine();
            ii++;
            //Default
            stbuild.Append("<BOMLine type=\"I\" sequencenumber=\"" + ii + "\">").AppendLine();
            stbuild.Append("<Item code=\"ETO\"/>").AppendLine();
            stbuild.Append(@"<Condition>S</Condition>").AppendLine();
            stbuild.Append(@"<Warehouse code=""KITR""/>").AppendLine();
            stbuild.Append(@"<BackFlush>0</BackFlush>").AppendLine();
            stbuild.Append(@"<EffectiveDate>" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "</EffectiveDate>").AppendLine();//10-05-2018
                                                                                                                            //stbuild.Append(@"<EffectiveDate>2000-01-10</EffectiveDate>").AppendLine();
                                                                                                                            //stbuild.Append(@"<ExpiryDate>2000-01-10</ExpiryDate>").AppendLine();//10-05-2018
            stbuild.Append(@"<Costcenter code=""MD1""/>").AppendLine();//30-04-2018
                                                                       //stbuild.Append(@"<CostCenter code=""001CC001""><Description>Default cost center</Description></CostCenter>").AppendLine();//30-04-2018
            stbuild.Append(@"<Quantity>1</Quantity>").AppendLine();
            stbuild.Append(@"</BOMLine>").AppendLine();
            ii++;

            stbuild.Append(@"</BOM>").AppendLine();
            stbuild.Append(@"</BOMs>").AppendLine();

            stbuild.Append(@"<FreeFields>").AppendLine();
            stbuild.Append(@"<FreeNumbers>").AppendLine();

            //stbuild.Append(@"<FreeText number = ""1""></FreeText>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""2""></FreeText>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""3""></FreeText>").AppendLine();
            if (totaltimerollup == "0")//22-05-2018
                stbuild.Append(@"<FreeNumber number=""4"">" + totaltimerollup + "</FreeNumber>").AppendLine();
            else
                stbuild.Append(@"<FreeNumber number=""4"">" + Math.Round(Convert.ToDouble(totaltimerollup) / 60, 2).ToString() + "</FreeNumber>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""5""></FreeText>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""6""></FreeText>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""7""></FreeText>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""8""></FreeText>").AppendLine();
            stbuild.Append(@"<FreeNumber number=""9"">" + SurfaceArea + "</FreeNumber>").AppendLine();
            //stbuild.Append(@"<FreeText number = ""10""></FreeText>").AppendLine();

            stbuild.Append(@"</FreeNumbers>").AppendLine();
            stbuild.Append(@"</FreeFields>").AppendLine();

            stbuild.Append(@"</Item>").AppendLine();
            stbuild.Append(@"</Items>").AppendLine();
            stbuild.Append(@"</eExact>").AppendLine();

            string st = stbuild.ToString();

            //System.IO.File.AppendAllText(@"C:\Users\BSethuraman\Downloads\BOM-test\test.xml", st);
            //System.Diagnostics.Debug.WriteLine(st);

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(st);

            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string xmlLocation = getpath.IniReadValue("FilePath", "XML_OUTPUT_LOCATION");

            string fulltitle = System.IO.Path.Combine(xmlLocation, "XML_" + drawingno + "_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + ".xml");

            try
            {
                if (System.IO.File.Exists(fulltitle))
                {
                    try
                    {
                        System.IO.File.Delete(fulltitle);
                    }
                    catch
                    {
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.Title = "Export Excel File To";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            xml.Save(saveFileDialog.FileName + ".xml");
                            //MessageBox.Show("XML created successfully");
                        }
                        else
                        {
                            MessageBox.Show("XML Creation Failed");
                        }
                        return;
                    }
                }

                xml.Save(fulltitle);

                // MessageBox.Show("XML created successfully");
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("XML output failed - " + ex.ToString());
            }

        }

        public void Generate(bool placeInAutoCAD = false, bool ExcelOutput = false, bool Timerollup = false, bool AutocadLayerwise = false)
        {
            //  System.Data.DataTable dt = DatabaseClass.GetProductsRM();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //List<KeyValuePair<string, string>> sheetvalue = new List<KeyValuePair<string, string>>();

            System.Data.DataTable dtnew = new System.Data.DataTable();
            if (dtnew == null)
            {
                System.Windows.MessageBox.Show("No items found in drawing");
                return;
            }

            //if (dtnew.Rows.Count > 0)
            //{

            System.Data.DataTable dtnewExcel = new System.Data.DataTable();

            //dtnewExcel.Columns.Add("Itemno");
            //dtnewExcel.Columns.Add("Quantity", System.Type.GetType("System.Int32"));
            //dtnewExcel.Columns.Add("Partno");
            //dtnewExcel.Columns.Add("Temp");
            //dtnewExcel.Columns.Add("Description");

            //foreach (DataRow item in dtnew.Rows)
            //{
            //    dtnewExcel.Rows.Add(item["Itemno"].ToString(), item["Quantity"].ToString(), item["Partno"].ToString(), "", item["Description"].ToString());
            //    //dtnewExcel.Rows.Add(item.item, item.TotalSum, item.part, "", item.desc);
            //}

            if (AutocadLayerwise)
            {

                dtnew = Get_Drawingdata(true, false);
                if (dtnew == null)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                if (dtnew.Rows.Count <= 0)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }

                DataView dv = dtnew.DefaultView;
                dv.Sort = "Layer,Itemno";
                dtnew = dv.ToTable();

                System.Data.DataTable layerlist = dv.ToTable(true, "Layer");

                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
                string BomLocation = getpath.IniReadValue("FilePath", "BOM_LOCATION");
                // int bom_rows = Convert.ToInt32(getpath.IniReadValue("FilePath", "BOM_ROWS"));
                foreach (DataRow datarow in layerlist.Rows)
                {

                    string layerval = (string)datarow[0];
                    System.Data.DataTable TowriteCad = dtnew.Select("Layer = '" + layerval + "'").CopyToDataTable();
                    dtnewExcel = TowriteCad.DefaultView.ToTable(false, "Itemno", "Quantity", "Partno", "Description");
                    try
                    {
                        /////create table on autocad
                        Point3d pt = Get_Max_Min_Point(layerval, BomLocation);
                        this.createautocadtablelayerwise(dtnewExcel, pt, layerval, BomLocation);//, bom_rows
                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.MessageBox.Show(ex.ToString());
                    }
                }
            }

            if (placeInAutoCAD)
            {
                dtnew = Get_Drawingdata(includeLenAttr: false);
                if (dtnew == null)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                if (dtnew.Rows.Count <= 0)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                dtnewExcel = dtnew.DefaultView.ToTable(false, "Itemno", "Quantity", "Partno", "Description");
                try
                {
                    /////create table on autocad                    
                    this.createautocadtable(dtnewExcel);
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.ToString());
                }
            }


            if (ExcelOutput)
            {
                // var description = getProjectDesc();
                dtnew = Get_Drawingdata(includeLenAttr: false);
                if (dtnew == null)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                if (dtnew.Rows.Count <= 0)
                {
                    MessageBox.Show("Drawing has no data");
                    return;
                }
                dtnewExcel = dtnew.DefaultView.ToTable(false, "Itemno", "Quantity", "Partno", "Temp", "Description");
                DataView dv = dtnewExcel.DefaultView;


                AlphanumComparator<string> comparer = new AlphanumComparator<string>();
                System.Data.DataTable dtNew = dv.Table.AsEnumerable().OrderBy(x => x.Field<string>("Partno"), comparer).CopyToDataTable();
                //dv.Sort = "Partno";
                dtnewExcel = dtNew;// dv.ToTable();
                foreach (DataRow dr in dtnewExcel.Rows)
                {
                    dr["PartNo"] = Regex.Replace(dr["PartNo"].ToString(), "[0-9]+", match => match.Value.PadLeft(9, '0'));
                }

                //MessageBox.Show("1");
                ///////Write on excel sheet
                ExcelCreationforXdata exc = new ExcelCreationforXdata();
                try
                {
                    string pN = Path.GetFileNameWithoutExtension(doc.Name).Replace("P_", "").Trim();
                    //MessageBox.Show("2");
                    exc.Writehead(dtnewExcel, System.IO.Path.GetFileNameWithoutExtension(doc.Name), sheetvalue, "", pN);
                    // System.Windows.MessageBox.Show("BOM Exported");
                }
                catch (System.Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.ToString());
                }
            }

            //}

        }

        public bool timerollupfunctionality(out string sumresult, out string dwgNumber, out string SurfaceAreaSum, bool writeexcel = true, bool readDesc = true, string projectDesc = "", bool divideBy60 = true, bool updateDWGAttr = true)
        {
            sumresult = "";
            dwgNumber = "";
            string revision = "";
            SurfaceAreaSum = "";
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //List<KeyValuePair<string, string>> sheetvalue = new List<KeyValuePair<string, string>>();

            System.Data.DataTable dtnew = Get_Drawingdata(updateDwgAttr: updateDWGAttr);
            if (dtnew == null)
            {
                System.Windows.MessageBox.Show("No items found in drawing");
                return false;
            }

            System.Data.DataTable dtnewExcel = new System.Data.DataTable();
            if (dtnew.Rows.Count > 0)
            {
                dtnewExcel = dtnew.DefaultView.ToTable(false, "Partno", "Description", "Timerollup", "SurfaceArea");

                System.Data.DataColumn newColumn = new System.Data.DataColumn("CP", typeof(System.String));
                newColumn.DefaultValue = "C";
                dtnewExcel.Columns.Add(newColumn);
                dtnewExcel.Columns["CP"].SetOrdinal(0);
            }

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
            string tempSum = (dtnewExcel.Rows.Count > 0) ? dtnewExcel.AsEnumerable().Sum(x => Convert.ToDouble(x["Timerollup"])).ToString() : "0";
            SurfaceAreaSum = (dtnewExcel.Rows.Count > 0) ? dtnewExcel.AsEnumerable().Sum(x => Convert.ToDouble(x["SurfaceArea"])).ToString() : "0";
            SurfaceAreaSum = Math.Round(Convert.ToDouble(SurfaceAreaSum), 3).ToString();
            //string dwgNumber = "";
            if (sheetvalue != null)
            {
                if (sheetvalue.Count > 0)
                {
                    foreach (KeyValuePair<string, string> kvp in sheetvalue)
                    {
                        if (kvp.Key == "DRG")
                        {
                            dwgNumber = kvp.Value;
                        }
                        else if (kvp.Key.Equals("REV", StringComparison.InvariantCultureIgnoreCase))
                        {
                            revision = kvp.Value;
                            //title = kvp.Value;
                        }
                    }
                }
            }

            View.Timerollup _frm = new View.Timerollup();
            _frm.totalTimeFromDwg = Convert.ToDouble(tempSum);
            _frm.drawingNumber = dwgNumber;
            System.Collections.ObjectModel.ObservableCollection<View.clsTimeRollUp> myCollection = new System.Collections.ObjectModel.ObservableCollection<View.clsTimeRollUp>(_col as List<View.clsTimeRollUp>);
            _frm.timeRollupCol = myCollection;
            bool? result = _frm.ShowDialog();
            //if (result != true)
            //  return false;

            if (dtnewExcel.Columns.Count <= 0)
            {
                dtnewExcel.Columns.Add("Partno");
                dtnewExcel.Columns.Add("Description");
                dtnewExcel.Columns.Add("Timerollup");
                //dtnewExcel.Columns.Add("Quantity");

                System.Data.DataColumn newColumn = new System.Data.DataColumn("CP", typeof(System.String));
                newColumn.DefaultValue = "C";
                dtnewExcel.Columns.Add(newColumn);
                dtnewExcel.Columns["CP"].SetOrdinal(0);
            }

            //Add only if yes given in message box in the form
            if (_frm.readresult == true)
            {
                _col = _frm.timeRollupCol.ToList();

                foreach (View.clsTimeRollUp _rollItem in _col.Where(a => a.include == true))
                {
                    dtnewExcel.Rows.Add("O", "", _rollItem.Category, _rollItem.Time);
                }

                tempSum = _frm.finalTimeFromUI.ToString();
            }
            else
                return false;



            string sumObject = (dtnewExcel.Rows.Count > 0) ? dtnewExcel.AsEnumerable().Sum(x => Convert.ToInt32(x["Timerollup"])).ToString() : "0";
            //dtnewExcel.Rows.Add("", "", "Total", sumObject);
            if (divideBy60)
                sumObject = (sumObject != "0") ? (Convert.ToDouble(sumObject) / 60).ToString() : sumObject;//Convert to mins
            try
            {
                sumresult = Math.Round(Convert.ToDouble(sumObject), 2).ToString();
            }
            catch { }
            dtnewExcel.Rows.Add("", "", "Total", sumresult);
            //DialogResult msgresult = MessageBox.Show("Do you want to continue with this time estimated \n Total time :" + sumObject, "", MessageBoxButtons.YesNo);

            //if (msgresult == DialogResult.Yes)
            //{
            ExcelCreationforXdata exc = new ExcelCreationforXdata();
            try
            {
                //dtnewExcel.AsEnumerable.ForEach(x => x["Timerollup"] = (double)x["Timerollup"] * (double)x["Quantity"]);
                //dtnewExcel = dtnewExcel.DefaultView.ToTable(false, "Partno", "Description", "Timerollup");


                //System.Data.DataTable dtResult = new System.Data.DataTable();
                //dtResult.Columns.Add("Partno");
                //dtResult.Columns.Add("Description");
                //dtResult.Columns.Add("Timerollup");

                //System.Data.DataColumn newColumn = new System.Data.DataColumn("CP", typeof(System.String));
                //newColumn.DefaultValue = "C";
                //dtResult.Columns.Add(newColumn);
                //dtResult.Columns["CP"].SetOrdinal(0);

                //foreach(DataRow drr in dtResult.Rows)
                //{
                //    dtResult.Rows.Add(drr["CP"].ToString(), drr["Partno"].ToString(), drr["Description"].ToString(), (Convert.ToDouble(drr["Timerollup"].ToString()) * Convert.ToDouble(drr["Quantity"].ToString())).ToString());
                //}
                if (writeexcel)
                {
                    if (readDesc)
                        projectDesc = getProjectDesc();
                    try
                    {
                        View.ProgressBar _pBar = new View.ProgressBar("Publish in progress");
                        _pBar.Show();
                        try
                        {// 15-05-2018, Added validations to avoid error message on no data
                            if (dtnewExcel != null)
                            {
                                if (dtnewExcel.Rows.Count > 0)
                                {
                                    if (dtnewExcel.AsEnumerable().Where(r => r.Field<string>("Timerollup") != "0").Any())
                                    {
                                        System.Data.DataTable table = dtnewExcel.AsEnumerable()
                                                                      .Where(r => r.Field<string>("Timerollup") != "0")
                                                                      .CopyToDataTable();
                                        dtnewExcel = table;
                                        // exc.WriteonTimerollup(dtnewExcel, dwgNumber, Math.Round(Convert.ToDouble(tempSum) / 60, 2).ToString(), projectDesc);//30-04-2018
                                        exc.WriteonTimerollup(dtnewExcel, dwgNumber, Math.Round(Convert.ToDouble(tempSum), 2).ToString(), projectDesc, revision: revision);
                                        _pBar.Close();
                                        return true;
                                    }
                                }
                            }
                        }
                        catch { }
                        _pBar.Close();
                    }
                    catch { }
                    MessageBox.Show("Time Rollup has no data to print", "SSDL - Time Rollup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }
            //}
            return true;
        }


        public string getProjectDesc()
        {
            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string maxChar = getpath.IniReadValue("FilePath", "MAX_BOM_CHAR");

            int maxCharAllowed = (string.IsNullOrEmpty(maxChar)) ? 0 : Convert.ToInt16(maxChar);
            View.Desc _desc = new View.Desc();
            _desc.maxAllowed = maxCharAllowed;
            _desc.ShowDialog();

            if (_desc.cancel == false)
                return _desc.textBoxdesc.Text;
            else
                return "";


            ///---Not reachable code

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor acDocEd = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            // Create a TypedValue array to define the filter criteria
            TypedValue[] acTypValAr = new TypedValue[2];
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 0);
            acTypValAr.SetValue(new TypedValue((int)DxfCode.Text, "*Label*"), 1);

            // Assign the filter criteria to a SelectionFilter object
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

            PromptEntityOptions pOts = new PromptEntityOptions("Select Description");
            pOts.AllowNone = true;
            pOts.SetRejectMessage("Select a MTEXT starts with Label");
            pOts.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.MText), true);
            // Request for objects to be selected in the drawing area
            PromptEntityResult acSSPrompt;
            acSSPrompt = acDocEd.GetEntity(pOts);

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    DBObject dbObj = tr.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead);
                    if (dbObj is MText)
                    {
                        string result = (dbObj as MText).Text.Replace("Label block", "").Replace("LABEL BLOCK", "").Replace(Environment.NewLine, " ").Replace("\n", " ").Trim();
                        RegexOptions options = RegexOptions.None;
                        Regex regex = new Regex("[ ]{2,}", options);
                        result = regex.Replace(result, " ");

                        return result;
                    }
                    else
                        return "";
                }

                //return acSSet.Replace("Label block", "").Trim();
            }
            else
            {
                return "";
            }
        }


        public void createautocadtablelayerwise(System.Data.DataTable dt, Point3d insertionpoint, string layername, string Bom_Location)//, int bom_rows,
        {
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No items found", "SSDL");
                return;
            }
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            //int totalRowCount = dt.Rows.Count;
            //int quotient = Convert.ToInt32(totalRowCount) / Convert.ToInt32(bom_rows);
            //int reminder = totalRowCount % bom_rows;
            //int ceiling = quotient + reminder;
            //int totalpagescount = (int)ceiling;
            // for (int loop = 1; loop <= ceiling; loop++)
            //{
            //if (loop > 1)
            //{
            //    insertionpoint = Get_Max_Min_Point(layername, Bom_Location);
            //}
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the blocktable, get the modelspace
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                /////creating table
                Table tb = new Table();
                tb.TableStyle = db.Tablestyle;
                tb.SetSize(dt.Rows.Count + 2, dt.Columns.Count);// dt.Rows.Count + 2, dt.Columns.Count);
                tb.SetRowHeight(3);
                tb.SetColumnWidth(20);

                tb.Columns[dt.Columns.Count - 1].Width = 70;

                tb.Layer = layername;
                tb.Position = insertionpoint;
                tb.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);

                string filepath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SheetFormat.dwg");
                ObjectId btrId = bt.GetBlock(filepath);

                Point3d ptInsert = new Point3d(0, 0, 0);
                Dictionary<string, string> attValues = new Dictionary<string, string>();
                BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                string blockName = System.IO.Path.GetFileNameWithoutExtension(filepath);

                int i = 1;
                int j = 0;
                foreach (System.Data.DataColumn dc in dt.Columns)
                {

                    tb.Cells[i, j].TextHeight = 2.5;
                    tb.Cells[i, j].TextString = dc.ColumnName.ToUpper();
                    tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
                    j++;
                }
                i++;
                int row_count = 1;
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    j = 0;
                    //if (row_count < loop * bom_rows && row_count > ((loop - 1) * bom_rows))
                    // {
                    foreach (Object dc in dr.ItemArray)
                    {
                        tb.Cells[i, j].TextHeight = 2.5;// tb.SetTextHeight(i, j, 1);
                        tb.Cells[i, j].TextString = dc.ToString();// tb.SetTextString(i, j, dc.ToString());
                        tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;// tb.SetAlignment(i, j, CellAlignment.MiddleCenter);
                        j++;
                    }
                    i++;
                    // }
                    row_count++;
                }
                tb.DeleteRows(0, 1);
                tb.GenerateLayout();

                btr.AppendEntity(tb);
                tr.AddNewlyCreatedDBObject(tb, true);
                tr.Commit();
            }
            //}
        }

        public void createautocadtable(System.Data.DataTable dt)
        {
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No items found", "SSDL");
                return;
            }
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the blocktable, get the modelspace
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                /////creating table
                Table tb = new Table();
                tb.TableStyle = db.Tablestyle;
                tb.SetSize(dt.Rows.Count + 2, dt.Columns.Count);
                tb.SetRowHeight(3);
                tb.SetColumnWidth(20);
                tb.Columns[dt.Columns.Count - 1].Width = 70;
                tb.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                //tb.Layer = layername;
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("");

                // Prompt for the start point
                pPtOpts.Message = "\n Pick Insertion point: ";
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                Point3d ptStart = pPtRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status != PromptStatus.OK) return;

                tb.Position = pPtRes.Value;// new Point3d(30, 260, 0);
                                           ////tb.SetCellType

                string filepath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SheetFormat.dwg");
                ObjectId btrId = bt.GetBlock(filepath);

                Point3d ptInsert = new Point3d(0, 0, 0);
                Dictionary<string, string> attValues = new Dictionary<string, string>();

                //attValues.Add(partnoatt, partnoval);            

                BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                string blockName = System.IO.Path.GetFileNameWithoutExtension(filepath);

                // BlockReference br = curSpace.InsertBlockReference(blockName, ptInsert, attValues);

                int i = 1;
                int j = 0;
                foreach (System.Data.DataColumn dc in dt.Columns)
                {

                    tb.Cells[i, j].TextHeight = 2.5;
                    tb.Cells[i, j].TextString = dc.ColumnName.ToUpper();
                    tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
                    j++;
                }


                i++;
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    j = 0;

                    foreach (Object dc in dr.ItemArray)
                    {
                        tb.Cells[i, j].TextHeight = 2.5;// tb.SetTextHeight(i, j, 1);
                        tb.Cells[i, j].TextString = dc.ToString();// tb.SetTextString(i, j, dc.ToString());
                        tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;// tb.SetAlignment(i, j, CellAlignment.MiddleCenter);
                        j++;
                    }

                    i++;
                }
                tb.DeleteRows(0, 1);
                tb.GenerateLayout();




                //var tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                //var romansID = tst.Has("Roman") ? tst["Roman"] : db.Textstyle;
                //var tsd = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
                //if (!tsd.Contains("SomeStyle"))
                //{
                //    var ts = new TableStyle();
                //    ts.SetTextHeight(5, (int)RowType.TitleRow);
                //    ts.SetTextHeight(3.5, (int)RowType.HeaderRow);
                //    ts.SetTextHeight(2, (int)RowType.DataRow);
                //    ts.SetTextStyle(romansID, (int)(RowType.HeaderRow | RowType.TitleRow | RowType.DataRow));
                //    ts.HorizontalCellMargin = 0.5;
                //    ts.VerticalCellMargin = 0.5;
                //    ts.SetGridColor(
                //    Color.FromColorIndex(ColorMethod.ByAci, 2),
                //    (int)GridLineType.AllGridLines, (int)RowType.DataRow);
                //    ts.SetGridColor(
                //    Color.FromColorIndex(ColorMethod.ByAci, 3),
                //    (int)GridLineType.AllGridLines, (int)RowType.HeaderRow);
                //    ts.SetGridColor(
                //    Color.FromColorIndex(ColorMethod.ByAci, 4),
                //    (int)GridLineType.AllGridLines, (int)RowType.TitleRow);
                //    tsd.UpgradeOpen();
                //    tsd.SetAt("SomeStyle", ts);
                //    tr.AddNewlyCreatedDBObject(ts, true);

                btr.AppendEntity(tb);
                tr.AddNewlyCreatedDBObject(tb, true);
                tr.Commit();
            }
        }


    }

    public class AlphanumComparator<T> : IComparer<T>
    {
        private enum ChunkType { Alphanumeric, Numeric };
        private bool InChunk(char ch, char otherCh)
        {
            ChunkType type = ChunkType.Alphanumeric;

            if (char.IsDigit(otherCh))
            {
                type = ChunkType.Numeric;
            }

            if ((type == ChunkType.Alphanumeric && char.IsDigit(ch))
                || (type == ChunkType.Numeric && !char.IsDigit(ch)))
            {
                return false;
            }

            return true;
        }

        public int Compare(T x, T y)
        {
            String s1 = x as string;
            String s2 = y as string;
            if (s1 == null || s2 == null)
            {
                return 0;
            }

            int thisMarker = 0, thisNumericChunk = 0;
            int thatMarker = 0, thatNumericChunk = 0;

            while ((thisMarker < s1.Length) || (thatMarker < s2.Length))
            {
                if (thisMarker >= s1.Length)
                {
                    return -1;
                }
                else if (thatMarker >= s2.Length)
                {
                    return 1;
                }
                char thisCh = s1[thisMarker];
                char thatCh = s2[thatMarker];

                StringBuilder thisChunk = new StringBuilder();
                StringBuilder thatChunk = new StringBuilder();

                while ((thisMarker < s1.Length) && (thisChunk.Length == 0 || InChunk(thisCh, thisChunk[0])))
                {
                    thisChunk.Append(thisCh);
                    thisMarker++;

                    if (thisMarker < s1.Length)
                    {
                        thisCh = s1[thisMarker];
                    }
                }

                while ((thatMarker < s2.Length) && (thatChunk.Length == 0 || InChunk(thatCh, thatChunk[0])))
                {
                    thatChunk.Append(thatCh);
                    thatMarker++;

                    if (thatMarker < s2.Length)
                    {
                        thatCh = s2[thatMarker];
                    }
                }

                int result = 0;
                // If both chunks contain numeric characters, sort them numerically
                if (char.IsDigit(thisChunk[0]) && char.IsDigit(thatChunk[0]))
                {
                    thisNumericChunk = Convert.ToInt32(thisChunk.ToString());
                    thatNumericChunk = Convert.ToInt32(thatChunk.ToString());

                    if (thisNumericChunk < thatNumericChunk)
                    {
                        result = -1;
                    }

                    if (thisNumericChunk > thatNumericChunk)
                    {
                        result = 1;
                    }
                }
                else
                {
                    result = thisChunk.ToString().CompareTo(thatChunk.ToString());
                }

                if (result != 0)
                {
                    return result;
                }
            }

            return 0;
        }


    }

}
