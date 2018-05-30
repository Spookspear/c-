#pragma warning disable IDE1006 // Naming Styles

// ask around how to add a comman class that I can share with this


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
// using ExcelRibbon = ToolbarOfFunctions.ExcelRibbon;

namespace ToolbarOfFunctions
    {
    public partial class ThisAddIn
    {
        internal readonly IntPtr Handle;
        //internal IntPtr Handle;
        //IntPtr Handle = Process.GetCurrentProcess().MainWindowHandle;
        // internal readonly IntPtr Handle = Process.GetCurrentProcess().MainWindowHandle;

        bool boolDisplayMessage;
        frmSettings myForm = new frmSettings();

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {

            // get the values from the form?
            boolDisplayMessage = myForm.chkProduceMessageBox.Checked;


        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        // Clear the workbook
        public void zapWorksheet(Excel.Workbook Wkb, int intFirstRow =2) {
            Excel.Worksheet Wks;
            Excel.Range xlCell;

            Wks = Wkb.ActiveSheet;

            // int intFirstRow = 2;
            int intLastRow = getLastRow(Wks);
            xlCell = Wks.get_Range("A" + intFirstRow + ":A" + intLastRow);

            if (Wks.Name != "InternalParameters") {
                if (intLastRow > intFirstRow) {
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);
                }
            } else {
                MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

        }

        internal void openSettingsForm(Excel.Workbook activeWorkbook)
        {

            // frmSettings myForm = new frmSettings();
            myForm.ShowDialog();

            /// ask nicola
            // ThisAddIn.btnDealWithSingleDuplicates.Label = "Hi";
            // ExcelRibbon ThisAddIn.btnDealWithSingleDuplicates.Label = "Hi";

            MsgBox("Hi");


            // throw new NotImplementedException();
        }

        public void MsgBox(string strMessage, string strWhichIcon = "Information") {
            MessageBoxIcon whichIcon = MessageBoxIcon.Information;
            string strCaption = strWhichIcon;

            switch (strWhichIcon)
            {
                case "Question":
                    whichIcon = MessageBoxIcon.Question;
                    break;

                case "Error":
                    whichIcon = MessageBoxIcon.Error;
                    break;

                case "Information":
                    whichIcon = MessageBoxIcon.Information;
                    break;

            }

            MessageBox.Show(strMessage, strCaption, MessageBoxButtons.OK, whichIcon);
            // MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question

        }


        public void readFolders(Excel.Workbook Wkb) {

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

            FolderBrowserDialogEx cfbd = new FolderBrowserDialogEx() {
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

            if (dlgReadExtraDetails != DialogResult.Cancel) {

                if (cfbd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK) {

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
                    Wks.Cells[1, (i+1)].value = strWords[i];
                
            Wks.Range["A1:E1"].Font.Bold = true;

        }

        public static DialogResult askGetExtraDetails() {
            DialogResult dlgResult = MessageBox.Show("Populate extract detail columns?", "Question", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            return dlgResult;
        }

        public static void getExtraDetails(string file, Excel.Worksheet Wks, int gintFileCount) {

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


            if (strFileVersioninfo != "0.0.0.0") { 
                Wks.Cells[gintFileCount, 4].value = strFileVersioninfo;
            }


            Wks.Cells[gintFileCount, 5].value = oFileInfo.Name;                                 // File Name Extracted
        }

        public static void directorySearch(string root, Excel.Worksheet Wks, int gintFileCount, bool boolExtraDetails, bool isRootItrated) {

            if (!isRootItrated) {
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
            if (subDirectories?.Any() == true) {
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
            boolDisplayMessage = myForm.chkProduceMessageBox.Checked;

            Excel.Worksheet Wks;   // get current sheet

            Wks = Wkb.ActiveSheet;

            int intStartColumToCheck = 1;
            string strColumnName = getExcelColumnName(intStartColumToCheck);

            DialogResult dlgResult = DialogResult.Yes;

            if (boolDisplayMessage)
            {
                dlgResult = MessageBox.Show("Duplicate Rows Check on column: " + strColumnName + " - worksheet name: " + Wks.Name, "Duplicate Rows Check", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            if (dlgResult == DialogResult.Yes)
            {
                int intStartRow = 1;

                int intLastRow = getLastRow(Wks);
                for (int intSourceRow = intStartRow; intSourceRow <= intLastRow; intSourceRow++)
                {
                    if (Wks.Cells[intSourceRow, intStartColumToCheck].Value == Wks.Cells[intSourceRow + 1, intStartColumToCheck].Value)
                        colourCells(Wks, (intSourceRow+1), "Error");
                
                }


                if (boolDisplayMessage)
                    MsgBox("Complete ...");

            }
        }


        public void compareSheets(Excel.Workbook Wkb) {

            Excel.Worksheet Wks1;   // get current sheet
            Excel.Worksheet Wks2;   // get sheet next door
            string strClearOrColour = "Colour";

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

                int intSheetLastRow1 = getLastRow(Wks1);
                int intSheetLastRow2 = getLastRow(Wks2);

                for (int intSourceRow = intStartRow; intSourceRow <= intSheetLastRow1; intSourceRow++)
                {
                    // read in vlaue from sheet 
                    // maybe I should ready all into arrayS?

                    strValue1 = Wks1.Cells[intSourceRow, intStartColumToCheck].Value;

                    intTargetRow = searchForValue(Wks2, strValue1, intStartColumToCheck);

                    if (intTargetRow > 0)
                    {
                        string stringCell1 = ""; string stringCell2 = "";

                        //  start from correct column
                        for (int intColCount  = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
                        {
                            if (!isEmptyCell(Wks1.Cells[intSourceRow, intColCount]))
                                stringCell1 = Wks1.Cells[intSourceRow, intColCount].Value.ToString();
                            
                            // need to handle nulls properly
                            if (!isEmptyCell(Wks2.Cells[intTargetRow, intColCount]))
                                stringCell2 = Wks2.Cells[intTargetRow, intColCount].Value.ToString();

                            if (stringCell1 == stringCell2)
                                intColScore++;
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

        public static int searchForValue(Excel.Worksheet Wks2, string searchString, int intStartColumToCheck)
        {

            Excel.Range colRange = Wks2.Columns["A:A"];             //get the range object where you want to search from

            Excel.Range resultRange = colRange.Find(

                    What: searchString,

                    LookIn: Excel.XlFindLookIn.xlValues,

                    LookAt: Excel.XlLookAt.xlPart,

                    // SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchOrder: Excel.XlSearchOrder.xlByColumns,

                    SearchDirection: Excel.XlSearchDirection.xlNext

                );                                                  // search searchString in the range, if find result, return a range
            /*
                if (resultRange is null) {
                    MessageBox.Show("Did not find " + searchString + " in column A");
                } else {
                    //then you could handle how to display the row to the label according to resultRange
                    MsgBox("found? - want to return the row no");
                }
                */

            return resultRange.Row;

        }


        public void deleteBlankLines(Excel.Workbook Wkb, string strMode) {
            Excel.Worksheet Wks;
            Wks = Wkb.ActiveSheet;

            string strResultsCol = "F";

            Stopwatch sw = new Stopwatch();

            // record times
            Wks.Range[strResultsCol + "1"].Value = DateTime.Now;

            sw.Start();

            if (Wks.Name != "InternalParameters")
            {
                if (strMode == "A") {
                    delLinesModeA(Wks);
                }

                if (strMode == "B") {
                    delLinesModeB(Wks);
                }

                if (strMode == "C") {
                    delLinesModeC(Wks);
                }


            }
            else {
                MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

            sw.Stop();

            Wks.Range[strResultsCol + "2"].Value = DateTime.Now;
            Wks.Range[strResultsCol + "3"].Value = sw.Elapsed.Milliseconds;
            MsgBox("Finshed ...");

        }


        private void delLinesModeA(Excel.Worksheet Wks) {
            Excel.Range xlCell;

            int intFirstRow = 2;
            int intColScore = 0;

            int intLastRow = getLastRow(Wks);
            int intLastCol = getLastCol(Wks);

            // loop along looking for data
            for (int intRows = intLastRow; intRows >= intFirstRow; intRows--) {
                Console.WriteLine(intRows);

                for (int intCols = 1; intCols <= intLastCol; intCols++) {
                    Console.WriteLine(intCols);

                    if (isEmptyCell(Wks.Cells[intRows, intCols]))
                        intColScore++;
                }

                if (intColScore == intLastCol) {
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
                int intRowLast = getLastRow(Wks);
                int intColLast = getLastCol(Wks);

                string strLastCol = getExcelColumnName(intColLast);

                int intRowToStartFrom = 1;
                int intColScore = 0;

                // loop along looking for data
                for (int intRows = 2; intRows <= intRowLast; intRows++)
                {
                    Console.WriteLine(intRows);

                    for (int intCols = 1; intCols <= intColLast; intCols++)
                    {
                        Console.WriteLine(intCols);

                        if (isEmptyCell(Wks.Cells[intRows, intCols]))
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
                MsgBox("There are no lines to delete", "Error");
                Console.WriteLine(excpt.Message);
            }
            

        }

        private void delLinesModeC(Excel.Worksheet worksheet)
        {
            // Excel.Application excel = new Excel.Application();

            // deleteEmptyRowsCols(worksheet);
            deleteEmptyRows(worksheet);

        }

        private static void deleteEmptyRows(Excel.Worksheet Wks)
        {
            Excel.Range xlTargetCells = Wks.UsedRange;
            object[,] allValues = (object[,])xlTargetCells.Cells.Value;
            int intTotalRows = xlTargetCells.Rows.Count;
            int intTotalCols = xlTargetCells.Columns.Count;

            List<int> lstEmptyRows = getEmptyRows(allValues, intTotalRows, intTotalCols);

            // now we have a list of the empty rows and columns we need to delete
            deleteRows(lstEmptyRows, Wks);
        }


        private static List<int> getEmptyRows(object[,] allValues, int intTotalRows, int intTotalCols)
        {
            List<int> lstEmptyRows = new List<int>();

            for (int i = 1; i < intTotalRows; i++)
            {
                if (isRowEmpty(allValues, i, intTotalCols))
                {
                    lstEmptyRows.Add(i);
                }
            }
            // sort the list from high to low
            return lstEmptyRows.OrderByDescending(x => x).ToList();
        }


        private static bool isRowEmpty(object[,] allValues, int rowIndex, int intTotalCols)
        {
            for (int i = 1; i < intTotalCols; i++)
            {
                if (allValues[rowIndex, i] != null)
                {
                    return false;
                }
            }
            return true;
        }

        private static void deleteRows(List<int> rowsToDelete, Excel.Worksheet worksheet)
        {
            // the rows are sorted high to low - so index's wont shift
            foreach (int rowIndex in rowsToDelete)
            {
                worksheet.Rows[rowIndex].Delete();
            }
        }


        // move these to common

        // I was after this 
        private string getExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        private bool isEmptyCell(Excel.Range xlCell) {
            bool boolRetVal = false;

            if (xlCell == null || xlCell.Value2 == null || xlCell.Value2.ToString() == "") {
                boolRetVal = true;
            }

            return boolRetVal;

        }

        public static int getLastCol(Excel.Worksheet Wks) {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
        }

        public static int getLastRow(Excel.Worksheet Wks) {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }
    
        private static void setCursorToWaiting() {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlWait;
        }
        private static void setCursorToDefault() {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlDefault;
        }

        private bool WorksheetExist(Excel.Workbook Wkb, string strSheetName) {
            bool found = false;

            foreach (Excel.Worksheet Wks in Wkb.Sheets) {

                MsgBox("inside WorksheetExist() - ws Name " + Wks.Name.ToLower());

                if (Wks.Name.ToLower() == strSheetName.ToLower()) {
                    found = true;
                    break;
                }
            }

            return found;
        }


        static void listSubFoldersAndFiles(string strSubFolderPath, Excel.Worksheet Wks, int gintFileCount) {
            // recursive function that will read from the current folder into selected wokrksheet
            try {
                foreach (string d in Directory.GetDirectories(strSubFolderPath)) {
                    foreach (string f in Directory.GetFiles(d)) {

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

        private static void graveYard()
        {
            // Excel.Worksheet activeWorksheet = null;
            // int cnt = 0;
            //foreach (Excel.Range element in range.Cells)            {
            //    if (element.Value2 != null)
            //    {
            //        cnt = cnt + 1;
            //    }
            //    System.Console.WriteLine(cnt);
            //}
            //MessageBox.Show(cnt.ToString());

            // Excel.Worksheet activeWorksheet;

            // MessageBox.Show(activeWorksheet.Name.ToLower());
            // string activeWorksheet = Wkb.Sheets[0].name;
            //activeWorksheet = Wkb.Sheets[1];
            //MessageBox.Show(activeWorksheet._CodeName);
            //WorksheetExist(Wkb, "Sheet1");

            // Excel.Range cell = activeWorksheet.get_Range(x.column + x.row);
            // string activeWorksheetName = activeWorksheet.Name;
            // MessageBox.Show(activeWorksheet.Name);
            // MessageBox.Show(activeWorksheetName);
            // setCursorToWaiting();

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
