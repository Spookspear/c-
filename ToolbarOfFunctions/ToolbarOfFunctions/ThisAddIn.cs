#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using DaveChambers.FolderBrowserDialogEx;

using System.ComponentModel;
using System.Data;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Tools.Ribbon;

using ToolbarOfFunctions_CommonClasses;


namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {
        internal readonly IntPtr Handle;

        public bool boolDisplayMessage;

        frmSettings frmSettings = new frmSettings();
        // frmSettings frmSettings = default(frmSettings);


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // get the values from the form?
            boolDisplayMessage = frmSettings.chkProduceMessageBox.Checked;


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        // Clear the workbook
        public void zapWorksheet(Excel.Workbook Wkb, int intFirstRow = 2)
        {
            Excel.Worksheet Wks;
            Excel.Range xlCell;

            Wks = Wkb.ActiveSheet;

            // int intFirstRow = 2;
            int intLastRow = CommonExcelClasses.getLastRow(Wks);
            xlCell = Wks.get_Range("A" + intFirstRow + ":A" + intLastRow);

            if (Wks.Name != "InternalParameters")
            {
                if (intLastRow > intFirstRow)
                {
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);
                }
            }
            else
            {
                CommonExcelClasses.MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

        }

        // internal void openSettingsForm(Excel.Workbook activeWorkbook)
        // {
        // frmSettings.ShowDialog();
        // 
        // }



        public void readFolders(Excel.Workbook Wkb)
        {

            Excel.Worksheet Wks;            // instantiate worksheet object
            Wks = Wkb.ActiveSheet;          // point to active sheet

            // custom extended class for browsing folders

            // am needing a folder full of files that will take a while to load
            // how many files?
            // 
            // C:\Temp\manyFiles-01
            string strPath;

            strPath = "C:\\Temp\\manyFiles-01";
            // strPath = "c:\\temp\\sfc"""

            FolderBrowserDialogEx cfbd = new FolderBrowserDialogEx()
            {
                Title = "Please Select Folder ...",
                SelectedPath = strPath,
                ShowEditbox = true,
                ShowNewFolderButton = false,
                RootFolder = Environment.SpecialFolder.Desktop,
                StartPosition = FormStartPosition.CenterScreen
            };

            // need a yes or no for reading in extra details
            DialogResult dlgReadExtraDetails = askGetExtraDetails();
            bool boolExtraDetails = false;

            if (dlgReadExtraDetails == DialogResult.Yes)
                boolExtraDetails = true;

            if (dlgReadExtraDetails == DialogResult.No)
                boolExtraDetails = false;

            // Excel.XlEnableCancelKey.xlInterrupt

            string strWhichColumn = "F";

            if (dlgReadExtraDetails != DialogResult.Cancel)
            {

                if (cfbd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {

                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    // Wks.Range[strWhichColumn + "1"].Value = DateTime.Now;

                    int gintFileCount = 2;

                    // zap the sheet before we start
                    zapWorksheet(Wkb, 1);

                    // string strPath = cfbd.SelectedPath;
                    directorySearch(cfbd.SelectedPath.ToString(), Wks, gintFileCount, boolExtraDetails, false);

                    writeHeaders(Wks, "FILES", boolExtraDetails);
                    Wks.Columns.AutoFit();

                    // Wks.Range[strWhichColumn + "2"].Value = DateTime.Now;
                    Wks.Range[strWhichColumn + "1"].Value = sw.Elapsed.Milliseconds;

                }
            }

            // MsgBox("Finished ...");

        }


        public void writeHeaders(Excel.Worksheet Wks, string strDoWhat, bool boolExtraDetails)
        {
            string strHead;

            if (boolExtraDetails)
                strHead = "File Name;Date Last Accessed;Size;Version;File Name Extracted;";
            else
                strHead = "FileName;;;;;";

            string[] strWords = strHead.Split(';');

            if (strDoWhat == "FILES")
                for (int i = 0; i <= strWords.GetUpperBound(0); i++)
                    Wks.Cells[1, (i + 1)].value = strWords[i];

            Wks.Range["A1:E1"].Font.Bold = true;

        }

        public static DialogResult askGetExtraDetails()
        {
            DialogResult dlgResult = MessageBox.Show("Populate extract detail columns?", "Question", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            return dlgResult;
        }

        public static void getExtraDetails(string file, Excel.Worksheet Wks, int gintFileCount)
        {

            FileInfo oFileInfo = new FileInfo(file);
            FileVersionInfo oFileVersionInfo = FileVersionInfo.GetVersionInfo(file);            // Get file version info LastWriteTime, LastAccessTime
            Wks.Cells[gintFileCount, 2].value = oFileInfo.LastAccessTime;                        // date
            Wks.Cells[gintFileCount, 3].value = oFileInfo.Length;                               // Size -- .ToString();

            // Wks.Cells[gintFileCount, 4].value = oFileVersionInfo.FileVersion;
            // 0.0.0.0
            // 16.21.0.0

            string strFileVersioninfo = oFileVersionInfo.FileMajorPart.ToString() + "." +
                                                oFileVersionInfo.FileMinorPart.ToString() + "." +
                                                oFileVersionInfo.FileBuildPart.ToString() + "." +
                                                oFileVersionInfo.FilePrivatePart.ToString();


            if (strFileVersioninfo != "0.0.0.0")
            {
                Wks.Cells[gintFileCount, 4].value = strFileVersioninfo;
            }


            Wks.Cells[gintFileCount, 5].value = oFileInfo.Name;                                 // File Name Extracted
        }

        public static void directorySearch(string root, Excel.Worksheet Wks, int gintFileCount, bool boolExtraDetails, bool isRootItrated)
        {

            if (!isRootItrated)
            {
                var rootDirectoryFiles = Directory.GetFiles(root);
                foreach (var file in rootDirectoryFiles)
                {

                    Console.WriteLine(file);
                    Wks.Cells[gintFileCount, 1].value = file.ToString();

                    if (boolExtraDetails)
                        getExtraDetails(file, Wks, gintFileCount);

                    gintFileCount++;
                }
            }

            // c# code to stop a macro running

            var subDirectories = Directory.GetDirectories(root);
            // does this need to be var?
            if (subDirectories?.Any() == true)
            {
                foreach (var directory in subDirectories)
                {
                    var files = Directory.GetFiles(directory);
                    foreach (var file in files)
                    {

                        Console.WriteLine(file);
                        Wks.Cells[gintFileCount, 1].value = file.ToString();
                        if (boolExtraDetails)
                            getExtraDetails(file, Wks, gintFileCount);

                        gintFileCount++;
                    }
                    directorySearch(directory, Wks, gintFileCount, boolExtraDetails, true);
                }
            }
        }


        internal void dealWithSingleDuplicates(Excel.Workbook Wkb)
        {

            // load relevant details from form
            boolDisplayMessage = frmSettings.chkProduceMessageBox.Checked;

            Excel.Worksheet Wks;   // get current sheet

            Wks = Wkb.ActiveSheet;

            int intStartColumToCheck = 1;
            string strColumnName = CommonExcelClasses.getExcelColumnName(intStartColumToCheck);

            DialogResult dlgResult = DialogResult.Yes;

            if (boolDisplayMessage)
            {
                dlgResult = MessageBox.Show("Duplicate Rows Check on column: " + strColumnName + " - worksheet name: " + Wks.Name, "Duplicate Rows Check", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            if (dlgResult == DialogResult.Yes)
            {
                int intStartRow = 1;

                int intLastRow = CommonExcelClasses.getLastRow(Wks);
                for (int intSourceRow = intStartRow; intSourceRow <= intLastRow; intSourceRow++)
                {
                    if (Wks.Cells[intSourceRow, intStartColumToCheck].Value == Wks.Cells[intSourceRow + 1, intStartColumToCheck].Value)
                        colourCells(Wks, (intSourceRow + 1), "Error");

                }


                if (boolDisplayMessage)
                    CommonExcelClasses.MsgBox("Complete ...");

            }
        }


        public void compareSheets(Excel.Workbook Wkb)
        {

            // this whoile thing needs to be in a try - 1gvb2

            Excel.Worksheet Wks1;   // get current sheet
            Excel.Worksheet Wks2;   // get sheet next door

            // string strClearOrColour = "Colour";     // need to get this from somewhere
            // strClearOrColour = (string)frmSettings.cmboCompareDifferences.SelectedItem;  // read from settings

            string strClearOrColour = SaveXML.readProperty("strCompareOrColour");    // read from function, that gets data from class

            /* 1gvb1 - can I read thisd from the object?
            InformationFromSettingsForm info = new InformationFromSettingsForm();
            strClearOrColour = info.HighLightOrDelete;
            */

            Wks1 = Wkb.ActiveSheet;
            Wks2 = Wkb.Sheets[Wks1.Index + 1];

            DialogResult dlgResult = MessageBox.Show("Compare: Worksheet: " + Wks1.Name + " against: " + Wks2.Name + " and " + strClearOrColour + " ones which are the same?", "Compare Sheets", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (dlgResult == DialogResult.Yes)
            {

                int intTargetRow = 0;
                int intStartRow = 2;

                // how may columns to check ?
                int intNoCheckCols = 5;             // for later loop?
                int intStartColumToCheck = 1;
                int intColScore = 0;

                string strValue1 = "";

                int intSheetLastRow1 = CommonExcelClasses.getLastRow(Wks1);
                int intSheetLastRow2 = CommonExcelClasses.getLastRow(Wks2);

                for (int intSourceRow = intStartRow; intSourceRow <= intSheetLastRow1; intSourceRow++)
                {
                    // read in vlaue from sheet 
                    // maybe I should ready all into arrays - maybe later?

                    strValue1 = Wks1.Cells[intSourceRow, intStartColumToCheck].Value;

                    intTargetRow = CommonExcelClasses.searchForValue(Wks2, strValue1, intStartColumToCheck);

                    if (intTargetRow > 0)
                    {
                        string stringCell1 = ""; string stringCell2 = "";

                        //  start from correct column
                        for (int intColCount = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
                        {

                            if (!CommonExcelClasses.isEmptyCell(Wks1.Cells[intSourceRow, intColCount]))
                                stringCell1 = Wks1.Cells[intSourceRow, intColCount].Value.ToString();

                            // need to handle nulls properly
                            if (!CommonExcelClasses.isEmptyCell(Wks2.Cells[intTargetRow, intColCount]))
                                stringCell2 = Wks2.Cells[intTargetRow, intColCount].Value.ToString();

                            if (stringCell1 == stringCell2)
                                intColScore++;
                        }

                    }

                    // Score system = if all the same then can blue it
                    if (intColScore == intNoCheckCols)
                        colourCells(Wks1, intSourceRow, "OK");
                    else
                        colourCells(Wks1, intSourceRow, "Error");

                    intColScore = 0;

                }
            }
        }

        public void colourCells(Excel.Worksheet Wks, int intSourceRow, string strDoWhat)
        {
            int intStartColumToCheck = 1;
            int intNoCheckCols = 5;

            for (int intColCount = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
            {
                if (strDoWhat == "OK")
                    Wks.Cells[intSourceRow, intColCount].Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Blue);

                if (strDoWhat == "Error")
                    Wks.Cells[intSourceRow, intColCount].Font.Color = ColorTranslator.ToOle(System.Drawing.Color.Red);
            }

        }



        public void deleteBlankLines(Excel.Workbook Wkb, string strMode)
        {
            Excel.Worksheet Wks;
            Wks = Wkb.ActiveSheet;

            string strResultsCol = "F";

            Stopwatch sw = new Stopwatch();

            // record times
            Wks.Range[strResultsCol + "1"].Value = DateTime.Now;

            sw.Start();

            if (Wks.Name != "InternalParameters")
            {
                if (strMode == "A")
                {
                    delLinesModeA(Wks);
                }

                if (strMode == "B")
                {
                    delLinesModeB(Wks);
                }

                if (strMode == "C")
                {
                    delLinesModeC(Wks);
                }

            }
            else
            {
                CommonExcelClasses.MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

            sw.Stop();

            Wks.Range[strResultsCol + "2"].Value = DateTime.Now;
            Wks.Range[strResultsCol + "3"].Value = sw.Elapsed.Milliseconds;
            CommonExcelClasses.MsgBox("Finshed ...");

        }


        private void delLinesModeA(Excel.Worksheet Wks)
        {
            Excel.Range xlCell;

            int intFirstRow = 2;
            int intColScore = 0;

            int intLastRow = CommonExcelClasses.getLastRow(Wks);
            int intLastCol = CommonExcelClasses.getLastCol(Wks);

            // loop along looking for data
            for (int intRows = intLastRow; intRows >= intFirstRow; intRows--)
            {
                Console.WriteLine(intRows);

                for (int intCols = 1; intCols <= intLastCol; intCols++)
                {
                    Console.WriteLine(intCols);

                    if (CommonExcelClasses.isEmptyCell(Wks.Cells[intRows, intCols]))
                        intColScore++;
                }

                if (intColScore == intLastCol)
                {
                    string strRange = "A" + intRows + ":A" + intRows;
                    xlCell = Wks.get_Range(strRange);
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);

                }

                // re initilise the score
                intColScore = 0;
            }
        }

        private void delLinesModeB(Excel.Worksheet Wks)
        {

            var range = Wks.UsedRange;

            try
            {
                range.SpecialCells(XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;

                range.SpecialCells(XlCellType.xlCellTypeVisible).Delete(XlDeleteShiftDirection.xlShiftUp);
                range.EntireRow.Hidden = false;

                Excel.Range xlCell;

                // int intRowFirst = 2;
                int intRowLast = CommonExcelClasses.getLastRow(Wks);
                int intColLast = CommonExcelClasses.getLastCol(Wks);

                string strLastCol = CommonExcelClasses.getExcelColumnName(intColLast);

                int intRowToStartFrom = 1;
                int intColScore = 0;

                // loop along looking for data
                for (int intRows = 2; intRows <= intRowLast; intRows++)
                {
                    Console.WriteLine(intRows);

                    for (int intCols = 1; intCols <= intColLast; intCols++)
                    {
                        Console.WriteLine(intCols);

                        if (CommonExcelClasses.isEmptyCell(Wks.Cells[intRows, intCols]))
                            intColScore++;
                    }

                    if (intColScore == intColLast)
                    {
                        intRowToStartFrom = intRows;
                        break;
                    }

                    // re initilise the score
                    intColScore = 0;
                }

                // ask weather to delete
                // create range to the end
                if (intRowToStartFrom <= intRowLast)
                {
                    intRowToStartFrom = (intRowToStartFrom + 3);
                    string strRange = "A" + intRowToStartFrom + ":" + strLastCol + intRowLast;
                    xlCell = Wks.get_Range(strRange);

                    // DialogResult dlgResult = MessageBox.Show("Carry out deleting rows? range is: " + strRange, "Question", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    // if (dlgResult == DialogResult.Yes)
                    // {
                    // and delete it
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);
                    this.Application.ActiveWorkbook.Save();

                    // }

                }

                range.SpecialCells(XlCellType.xlCellTypeConstants).EntireRow.Hidden = false;

            }
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("There are no lines to delete", "Error");
                Console.WriteLine(excpt.Message);
            }


        }

        private void delLinesModeC(Excel.Worksheet worksheet)
        {
            // Excel.Application excel = new Excel.Application();

            // deleteEmptyRowsCols(worksheet);
            CommonExcelClasses.deleteEmptyRows(worksheet);

        }


        static void listSubFoldersAndFiles(string strSubFolderPath, Excel.Worksheet Wks, int gintFileCount)
        {
            // recursive function that will read from the current folder into selected wokrksheet
            try
            {
                foreach (string d in Directory.GetDirectories(strSubFolderPath))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {

                        Console.WriteLine(f);

                        Wks.Cells[gintFileCount, 1].value = f.ToString();
                        gintFileCount++;
                    }
                    listSubFoldersAndFiles(d, Wks, gintFileCount);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }

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
