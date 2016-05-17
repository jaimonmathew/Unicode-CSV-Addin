/*
    Copyright 2011 Jaimon Mathew www.jaimon.co.uk

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.     
 
*/

using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

namespace UnicodeCSVAddin
{
    public partial class ThisAddIn
    {
        private Excel.Application app;
        private List<string> unicodeFiles; //a list of opened Unicode CSV files. We populate this list on WorkBookOpen event to avoid checking for CSV files on every Save event.
        private bool sFlag = false;

        //Unicode file byte order marks.
        private const string UTF_16BE_BOM = "FEFF";
        private const string UTF_16LE_BOM = "FFFE";
        private const string UTF_8_BOM = "EFBBBF";

        private string lSeparator = ",";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = this.Application;
            unicodeFiles = new List<string>();
            app.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen);
            app.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(app_WorkbookBeforeClose);
            app.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(app_WorkbookBeforeSave);
            try
            {
                lSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            }
            catch
            {
                lSeparator = ",";
            }
        }

        void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            app = null;
            unicodeFiles = null;
        }

        void app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            //Override Save behaviour for Unicode CSV files.
            if (!SaveAsUI && !sFlag && unicodeFiles.Contains(Wb.FullName))
            {
                Cancel = true;
                SaveAsUnicodeCSV(false, false);
            }
            sFlag = false;
        }

        //This is required to show our custom Ribbon
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        void app_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            unicodeFiles.Remove(Wb.FullName);
            app.StatusBar = "Ready";
        }

        void app_WorkbookOpen(Excel.Workbook Wb)
        {
            //Check to see if the opened document is a Unicode CSV files, so we can override Excel's Save method
            if (Wb.FullName.ToLower().EndsWith(".csv") && isFileUnicode(Wb.FullName))
            {
                if (!unicodeFiles.Contains(Wb.FullName))
                {
                    unicodeFiles.Add(Wb.FullName);
                }
                app.StatusBar = Wb.Name + " has been opened as a Unicode CSV file";
            }
            else
            {
                app.StatusBar = "Ready";
            }
        }

        /// <summary>
        /// This method check whether Excel is in Cell Editing mode or not
        /// There are few ways to check this (eg. check to see if a standard menu item is disabled etc.)
        /// I know in cell editing mode app.DisplayAlerts throws an Exception, so here I'm relying on that behaviour
        /// </summary>
        /// <returns>true if Excel is in cell editing mode</returns>
        private bool isInCellEditingMode()
        {
            bool flag = false;
            try
            {
                app.DisplayAlerts = false; //This will throw an Exception if Excel is in Cell Editing Mode
            }
            catch (Exception)
            {
                flag = true;
            }
            return flag;
        }
        /// <summary>
        /// This will create a temporary file in Unicode text (*.txt) format, overwrite the current loaded file by replaing all tabs with a comma and reload the file.
        /// </summary>
        /// <param name="force">To force save the current file as a Unicode CSV.
        /// When called from the Ribbon items Save/SaveAs, <i>force</i> will be true
        /// If this parameter is true and the file name extention is not .csv, then a SaveAs dialog will be displayed to choose a .csv file</param>
        /// <param name="newFile">To show a SaveAs dialog box to select a new file name
        /// This will be set to true when called from the Ribbon item SaveAs</param>
        public void SaveAsUnicodeCSV(bool force, bool newFile)
        {
            app.StatusBar = "";
            bool currDispAlert = app.DisplayAlerts;
            bool flag = true;
            int i;
            string filename = app.ActiveWorkbook.FullName;

            if (force) //then make sure a csv file name is selected.
            {
                if (newFile || !filename.ToLower().EndsWith(".csv"))
                {
                    Office.FileDialog d = app.get_FileDialog(Office.MsoFileDialogType.msoFileDialogSaveAs);
                    i = app.ActiveWorkbook.Name.LastIndexOf(".");
                    if (i >= 0)
                    {
                        d.InitialFileName = app.ActiveWorkbook.Name.Substring(0, i);
                    }
                    else
                    {
                        d.InitialFileName = app.ActiveWorkbook.Name;
                    }
                    d.AllowMultiSelect = false;
                    Office.FileDialogFilters f = d.Filters;
                    for (i = 1; i <= f.Count; i++)
                    {
                        if ("*.csv".Equals(f.Item(i).Extensions))
                        {
                            d.FilterIndex = i;
                            break;
                        }
                    }
                    if (d.Show() == 0) //User cancelled the dialog
                    {
                        flag = false;
                    }
                    else
                    {
                        filename = d.SelectedItems.Item(1);
                    }
                }
                if (flag && !filename.ToLower().EndsWith(".csv"))
                {
                    MessageBox.Show("Please select a CSV file name first");
                    flag = false;
                }
            }

            if (flag && filename.ToLower().EndsWith(".csv") && (force || unicodeFiles.Contains(filename)))
            {
                if (isInCellEditingMode())
                {
                    MessageBox.Show("Please finish editing before saving");
                }
                else
                {
                    try
                    {
                        //Getting current selection to restore the current cell selection
                        Excel.Range rng = (Excel.Range)app.ActiveCell;
                        int row = rng.Row;
                        int col = rng.Column;

                        string tempFile = System.IO.Path.GetTempFileName();

                        try
                        {
                            sFlag = true; //This is to prevent this method getting called again from app_WorkbookBeforeSave event caused by the next SaveAs call
                            app.ActiveWorkbook.SaveAs(tempFile, Excel.XlFileFormat.xlUnicodeText);
                            app.ActiveWorkbook.Close();

                            if (new FileInfo(tempFile).Length <= (1024 * 1024)) //If its less than 1MB, load the whole data to memory for character replacement
                            {
                                File.WriteAllText(filename, File.ReadAllText(tempFile, Encoding.Unicode).Replace("\t", lSeparator), Encoding.Unicode);
                            }
                            else //otherwise read chunks for data (in 10KB chunks) into memory
                            {
                                using (StreamReader sr = new StreamReader(tempFile, true))
                                using (StreamWriter sw = new StreamWriter(filename, false, Encoding.Unicode))
                                {
                                    char[] buffer = new char[100 * 1024]; //100KB Chunks
                                    while (!sr.EndOfStream)
                                    {
                                        int cnt = sr.ReadBlock(buffer, 0, buffer.Length);
                                        for (i = 0; i < cnt; i++)
                                        {
                                            if (buffer[i] == '\t')
                                            {
                                                buffer[i] = lSeparator[0];
                                            }
                                        }
                                        sw.Write(buffer, 0, cnt);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            File.Delete(tempFile);
                        }

                        app.Workbooks.Open(filename, Type.Missing, Type.Missing, 6, Type.Missing, Type.Missing, Type.Missing, Type.Missing, lSeparator);
                        Excel.Worksheet ws = app.ActiveWorkbook.ActiveSheet;
                        ws.Cells[row, col].Select();
                        app.StatusBar = Path.GetFileName(filename) + " has been saved as a Unicode CSV";
                        if (!unicodeFiles.Contains(filename))
                        {
                            unicodeFiles.Add(filename);
                        }
                        app.ActiveWorkbook.Saved = true;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Error occured while trying to save this file as Unicode CSV: " + e.Message);
                    }
                    finally
                    {
                        app.DisplayAlerts = currDispAlert;
                    }
                }
            }
        }

        /// <summary>
        /// This method will try and read the first few bytes to see if it contains a Unicode BOM
        /// </summary>
        /// <param name="filename">File to check for including full path</param>
        /// <returns>true if its a Unicode file</returns>
        private bool isFileUnicode(string filename)
        {
            bool ret = false;
            try
            {
                byte[] buff = new byte[3];
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Read(buff, 0, 3);
                }

                string hx = "";
                foreach (byte letter in buff)
                {
                    hx += string.Format("{0:X2}", Convert.ToInt32(letter));
                    //Checking to see the first bytes matches with any of the defined Unicode BOM
                    //We only check for UTF8 and UTF16 here.
                    ret = UTF_16BE_BOM.Equals(hx) || UTF_16LE_BOM.Equals(hx) || UTF_8_BOM.Equals(hx);
                    if (ret)
                    {
                        break;
                    }
                }
            }
            catch (IOException)
            {
                //ignore any exception
            }
            return ret;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
