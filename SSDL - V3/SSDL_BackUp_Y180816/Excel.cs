using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;

namespace SSDL
{
    public class ExcelCreationforXdata
    {
        private static Microsoft.Office.Interop.Excel.Application mExcelApp;
        private Workbook mExportDoc;
        private Worksheet SheetName = default(Worksheet);

        public void Startapplication()
        {
            //if (mExcelApp == null)
            {
                try
                {
                    mExcelApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    mExcelApp = new Microsoft.Office.Interop.Excel.Application();
                }
                mExcelApp.Visible = false;
                // mExcelApp.WindowState = XlWindowState.xlMaximized;
            }
        }

        private bool IsOpened(string wbook)
        {
            bool isOpened = true;
            try
            {
                mExcelApp.Workbooks.get_Item(wbook);

            }
            catch (Exception)
            {
                isOpened = false;
            }
            return isOpened;
        }

        public void Writehead(System.Data.DataTable dt, string title, List<KeyValuePair<string, string>> titleblockvalue = null, string desc = "", string fileName = "")
        {
            string fulltitle = "";
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No items found", "SSDL");
                return;
            }
            this.Startapplication();
            if (this.IsOpened("BOM_TEMPLATE.xlsx"))
            {
                mExcelApp.Workbooks["BOM_TEMPLATE.xlsx"].Activate();
                mExcelApp.ActiveWorkbook.Close();
            }
            mExportDoc = mExcelApp.Workbooks.Open(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "BOM_TEMPLATE.xlsx"));
            if (mExportDoc != null)
            {
                SheetName = (Worksheet)mExportDoc.ActiveSheet;
                mExcelApp.Visible = false;
                //SheetName.Range["A1"].EntireRow.RowHeight = 75;
                //SheetName.Range["A1"].EntireRow.WrapText = true;

                //int i = 1;
                //foreach (DataColumn dc in dt.Columns)
                //{
                //    writeonheadcell(i, 1, dc.ColumnName.ToUpper());
                //    i++;
                //}

                //Changed on 09-02-2018
                // writeoncell(1, 2, desc);
                if (fileName.Length >= 9)
                    writeoncell(3, 4, fileName.Substring(0, 9));//22-05-2018
                else
                    writeoncell(3, 4, fileName);

                int i = 7;
                foreach (DataRow dr in dt.Rows)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        writeoncell(j + 1, i, dr[j].ToString(), addBorder: true, isNumber: (j == 1) ? true : false);
                    }
                    SheetName.Range[SheetName.Cells[i, 6], SheetName.Cells[i, 9]].Merge();
                    SheetName.Range[SheetName.Cells[i, 10], SheetName.Cells[i, 18]].Merge();
                    i++;
                }
                SheetName.Columns.AutoFit();

                //Page No - Excel Sheet Count
                try
                {
                    writeoncell(18, 4, mExportDoc.Sheets.Count.ToString());
                }
                catch { }

                if (titleblockvalue != null)
                {
                    if (titleblockvalue.Count > 0)
                    {
                        foreach (KeyValuePair<string, string> kvp in titleblockvalue)
                        {
                            if (kvp.Key.Equals("ECO", StringComparison.InvariantCultureIgnoreCase))
                            {
                                writeoncell(11, 4, kvp.Value);
                            }
                            // else if (kvp.Key.Equals("SHT", StringComparison.InvariantCultureIgnoreCase))
                            //{
                            // writeoncell(18, 4, kvp.Value);
                            //writeoncell(12, 8, kvp.Value);
                            //}
                            else if (kvp.Key.Equals("DRG", StringComparison.InvariantCultureIgnoreCase))
                            {
                                // writeoncell(12, 3, kvp.Value);
                                title = kvp.Value;
                            }
                            else if (kvp.Key.Equals("REV", StringComparison.InvariantCultureIgnoreCase))
                            {
                                writeoncell(7, 4, kvp.Value);
                                //title = kvp.Value;
                            }
                        }
                    }
                }

                IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
                string path = getpath.IniReadValue("FilePath", "BOM_OUTPUT_LOCATION");

                fulltitle = System.IO.Path.Combine(path, "BOM_" + title + "_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + ".xlsx");

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
                                mExportDoc.SaveAs(saveFileDialog.FileName + ".xlsx", ReadOnlyRecommended: true); //15-05-2018 - Added Read only
                                // No Need to close
                                mExportDoc.Close();
                                mExcelApp.Quit();
                                System.Diagnostics.Process.Start(saveFileDialog.FileName + ".xlsx");
                                return;
                            }
                        }
                    }
                    mExportDoc.SaveAs(fulltitle, ReadOnlyRecommended: true); //15-05-2018 - Added Read only
                    //No Need to close
                    mExportDoc.Close();
                    mExcelApp.Quit();
                    System.Diagnostics.Process.Start(fulltitle);
                }
                catch (System.Runtime.InteropServices.COMException exx)
                {
                    System.Windows.Forms.MessageBox.Show(fulltitle + " is already open, you cannot create another Excel BOM of same name");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Excel output failed - " + ex.ToString());
                }

            }
        }

        public void WriteonTimerollup(System.Data.DataTable dt, string title, string totalTime = "0", string desc = "", string revision = "")
        {
            this.Startapplication();

            if (this.IsOpened("TIME_ROLLUP_TEMPLATE.xlsx"))
            {
                mExcelApp.Workbooks["TIME_ROLLUP_TEMPLATE.xlsx"].Activate();
                mExcelApp.ActiveWorkbook.Close();
            }
            mExportDoc = mExcelApp.Workbooks.Open(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TIME_ROLLUP_TEMPLATE.xlsx"));

            if (mExportDoc != null)
            {
                SheetName = (Worksheet)mExportDoc.ActiveSheet;
            }

            int i = 8;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 7 + dt.Rows.Count)
                    break;
                for (int j = 0; j < dt.Columns.Count - 1; j++) //-1 added to avoid SurfaceArea Column
                {
                    writeoncell(j + 2, i, dr[j].ToString(), isNumber: (j == 3) ? true : false);
                }
                i++;
            }

            //object sumObject = dt.Compute("Sum(Timerollup)", "");

            ////15-5-2018, added to change to mins
            //if (totalTime != "0")
            //{
            //    totalTime = (Math.Round(Convert.ToDouble(totalTime) / 60, 2)).ToString();
            //}
            string sumObject = totalTime;// dt.AsEnumerable().Sum(x => Convert.ToInt32(x["Timerollup"]) / 2).ToString();
            writeoncell(5, 5, revision);//15-05-2018, Added revision to Time roll up output
            writeoncell(5, 6, sumObject.ToString());

            //Product Code
            writeoncell(3, 5, title.Replace("P_", ""), true);
            //Product Desc
            writeoncell(3, 4, desc, true);

            IniFile getpath = new IniFile(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.ini"));
            string path = getpath.IniReadValue("FilePath", "TIME_ROLLUP_OUTPUT_LOCATION");

            string fulltitle = System.IO.Path.Combine(path, "TIME_ROLLUP_" + title + "_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + ".xlsx");

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
                            fulltitle = saveFileDialog.FileName + ".xlsx";
                            mExportDoc.SaveAs(fulltitle, ReadOnlyRecommended: true); //15-05-2018 - Added Read only
                            //No need to close
                            mExportDoc.Close();
                            mExcelApp.Quit();
                            System.Diagnostics.Process.Start(fulltitle);
                            return;
                        }
                    }
                }
                mExportDoc.SaveAs(fulltitle, ReadOnlyRecommended: true); //15-05-2018 - Added Read only
                //No need to close
                mExportDoc.Close();
                mExcelApp.Quit();

                System.Diagnostics.Process.Start(fulltitle);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Excel output failed - " + ex.ToString());
            }

        }


        public void writeonheadcell(int col, int row, string towrite, bool isNumber = false)
        {
            string columnstring = getcolumnstring(col);

            Range Getrange = SheetName.Range[string.Concat(columnstring, row.ToString()), Missing.Value];
            Getrange.Value = towrite;

            Getrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            Getrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Getrange.Interior.Color = XlRgbColor.rgbLightGrey;
            Getrange.Font.Bold = true;
            Getrange.ColumnWidth = 20;
            Getrange.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

        }

        public void writeoncell(int col, int row, string towrite, bool textbold = false, bool addBorder = false, bool isNumber = false)
        {
            string columnstring = getcolumnstring(col);

            Range Getrange = SheetName.Range[string.Concat(columnstring, row.ToString()), Missing.Value];
            if (!isNumber)
                //    Getrange.NumberFormat = "#";
                //else
                Getrange.NumberFormat = "@";
            Getrange.Value = towrite;

            Getrange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            Getrange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            if (addBorder)
            {
                for (int i = 1; i <= 18; i++)
                {
                    Range Getrange1 = SheetName.Range[string.Concat(getcolumnstring(i), row.ToString()), Missing.Value];
                    Getrange1.Borders.Color = System.Drawing.Color.Black.ToArgb();
                }
            }

            if (textbold == true)
                Getrange.Font.Bold = true;
        }

        public string getcolumnstring(int columnvalue)
        {
            string columnstring;
            if (columnvalue < 27)
            {
                columnstring = this.Converttostring(columnvalue);
            }
            else
            {
                int firststring = columnvalue / 26;
                int secondstring = columnvalue % 26;
                columnstring = string.Concat(this.Converttostring(firststring), this.Converttostring(secondstring));
            }
            return columnstring;
        }

        public string Converttostring(int unicode)
        {
            char character = (char)(unicode + 64);
            return character.ToString();
        }

        public Range Mergecells(Worksheet worksheet, int col1, int row1, int col2, int row2)
        {
            Range Getrange = worksheet.Range[string.Concat(getcolumnstring(col1), row1.ToString()), string.Concat(getcolumnstring(col2), row2.ToString())];
            Getrange.Merge();
            return Getrange;
        }
    }
}
