﻿#pragma warning disable IDE1006 // Naming Styles

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
using ToolbarOfFunctions_MyConstants;
using System.Runtime.InteropServices;

// using System.Data.SqlTypes;

using System.DirectoryServices;

namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {
        internal readonly IntPtr Handle;

        public string LF = MyConstants._LF;

        public int C_COL_CATEGORY = 17;
        public int C_COL_TOTAL = 19;
        public int C_COL_DATE = 14;
        

        frmSettings frmSettings = new frmSettings();
        InformationForSettingsForm myData = new InformationForSettingsForm();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // myData = SaveXML.LoadData();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }





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

                    DateTime dteStart = DateTime.Now;

                    int gintFileCount = 2;

                    // zap the sheet before we start
                    CommonExcelClasses.zapWorksheet(Wks, 1);

                    // string strPath = cfbd.SelectedPath;
                    directorySearch(cfbd.SelectedPath.ToString(), Wks, gintFileCount, boolExtraDetails, false);

                    writeHeaders(Wks, "FILES", boolExtraDetails);
                    

                    DateTime dteEnd = DateTime.Now;
                    int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                    Wks.Range[strWhichColumn + "1"].Value = milliSeconds + " milliSeconds";

                }
            }

            // MsgBox("Finished ...");

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="Wks"></param>
        /// <param name="strDoWhat"></param>
        /// <param name="boolExtraDetails"></param>
        public void writeHeaders(Excel.Worksheet Wks, string strDoWhat, bool boolExtraDetails)
        {
            string strHead = "";

            if (strDoWhat == "FILES")
            {
                if (boolExtraDetails)
                    strHead = "File Name;Date Last Accessed;Size;Version;File Name Extracted;";
                else
                    strHead = "FileName;;;;;";

            }
            else if (strDoWhat == "ADGroups")
            {
                strHead = "Group Name;Group Description;IsSecurityGroup;Scope";
            }
            else if (strDoWhat == "ADUsers")
            {
                strHead = "Name;Full Name;Description;AccountDisabled";
            }

            string[] strWords = strHead.Split(';');

            for (int i = 0; i <= strWords.GetUpperBound(0); i++)
                    Wks.Cells[1, (i + 1)].value = strWords[i];

            Wks.Range["A1:E1"].Font.Bold = true;
            Wks.Columns.AutoFit();

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
                Wks.Cells[gintFileCount, 4].value = strFileVersioninfo;

            Wks.Cells[gintFileCount, 5].value = oFileInfo.Name;                                 // File Name Extracted
        }


        public static void directorySearch(string root, Excel.Worksheet Wks, int gintFileCount, bool boolExtraDetails, bool isRootItrated)
        {

            if (!isRootItrated)
            {
                var rootDirectoryFiles = Directory.GetFiles(root);
                foreach (var file in rootDirectoryFiles)
                {
                    // Console.WriteLine(file);
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


        /// <summary>
        /// dealWithSingleDuplicates
        /// Loops down a single column looking for duplicates
        /// </summary>
        internal void dealWithSingleDuplicates(Excel.Application xls)
        {

            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            string strColourOrDelete = myData.ColourOrDelete;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolClearFormatting = myData.ClearFormatting;

            // colours for the Colour or delete option
            Color clrFoundForeColour = ColorTranslator.FromHtml(myData.ColourFore_Found);
            Color clrFoundBackColour = ColorTranslator.FromHtml(myData.ColourBack_Found);

            decimal intStartRow = myData.ComparingStartRow;
            decimal intStartColumToCheck = myData.DupliateColumnToCheck;

            #endregion

            try
            {

                #region [Declare and instantiate variables for worksheet/book]
                Excel.Workbook Wkb = xls.ActiveWorkbook;
                Excel.Worksheet Wks;   // get current sheet

                Wks = Wkb.ActiveSheet;

                string strColumnName = CommonExcelClasses.getExcelColumnLetter((int)intStartColumToCheck);

                DialogResult dlgResult = DialogResult.Yes;

                string strMessage;

                int intLastRow = CommonExcelClasses.getLastRow(Wks);

                // start of loop
                decimal intSourceRow = intStartRow;
                #endregion 

                #region [Display a Message?]
                if (boolDisplayInitialMessage)
                {
                    strMessage = "";
                    strMessage = strMessage + "Worksheet: " + Wks.Name + LF;
                    strMessage = strMessage + "Column: " + strColumnName + LF; 
                    strMessage = strMessage + "and: " + strColourOrDelete + " ones which are the same";

                    if (booltimeTaken)
                    {
                        strMessage = strMessage + LF + " and display the time taken";
                    }

                    strMessage = strMessage + "?";

                    dlgResult = MessageBox.Show(strMessage, "Duplicate Rows Check", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    // remove formatting - format black and white but only if no was selected
                    if (dlgResult == DialogResult.No)
                    {
                        if (boolClearFormatting) 
                            CommonExcelClasses.clearFormattingRange(Wks);
                        
                    }

                }
                #endregion

                #region [Start of work]
                if (dlgResult == DialogResult.Yes)
                {

                    DateTime dteStart = DateTime.Now;

                    decimal curNoRecords = 0;


                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                    #region [Start of loop]
                    while (!CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow, intStartColumToCheck]))
                    {
                        // hightlight, delete or clear?
                        if (Wks.Cells[intSourceRow, intStartColumToCheck].Value == Wks.Cells[intSourceRow + 1, intStartColumToCheck].Value )
                        {
                            while (Wks.Cells[intSourceRow, intStartColumToCheck].Value == Wks.Cells[intSourceRow + 1, intStartColumToCheck].Value)
                            {
                                if (strColourOrDelete == "Colour")
                                {
                                    CommonExcelClasses.colourCells(Wks, (intSourceRow + 1), "Error", 1, clrFoundForeColour, clrFoundBackColour, false);
                                    intSourceRow++;
                                }
                                else if (strColourOrDelete == "Delete")
                                {
                                    Wks.Rows[intSourceRow].Delete();
                                } else {
                                    CommonExcelClasses.colourCells(Wks, (intSourceRow ), strColourOrDelete, 1, clrFoundForeColour, clrFoundBackColour, false);
                                    intSourceRow++;
                                }

                                curNoRecords++;

                                if (CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow+1, intStartColumToCheck]))
                                    break;                              

                            }
                        
                        }

                        intSourceRow++;
                    }
                    #endregion [Start of loop]

                    

                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);

                    #region [Display Complete Message]
                    if (boolDisplayCompleteMessage)
                    {

                        strMessage = "Complete ...";

                        if (booltimeTaken)
                        {
                            DateTime dteEnd = DateTime.Now;
                            int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                            strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds + LF;
                            strMessage = strMessage + LF + "And handled: " + curNoRecords.ToString() + " duplicates";

                        }
                        CommonExcelClasses.MsgBox(strMessage);
                    }
                    #endregion

                }
                #endregion [Start of work]

                #region [Release memory]
                Marshal.ReleaseComObject(Wks);
                Marshal.ReleaseComObject(Wkb);
                #endregion

            }
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("There was an error?", "Error");
                Console.WriteLine(excpt.Message);
            }
        }


        /// <summary>
        /// dealWithManyDuplicates
        /// Loops down many columns looking for duplicates
        /// </summary>
        internal void dealWithManyDuplicates(Excel.Application xls)
        {
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            string strColourOrDelete = myData.ColourOrDelete;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolClearFormatting = myData.ClearFormatting;

            // colours for the Colour or delete option
            Color clrFoundForeColour = ColorTranslator.FromHtml(myData.ColourFore_Found);
            Color clrFoundBackColour = ColorTranslator.FromHtml(myData.ColourBack_Found);

            decimal intStartRow = myData.ComparingStartRow;
            decimal intColumToCheck = myData.DupliateColumnToCheck;
            decimal intNoCheckCols = myData.NoOfColumnsToCheck;          // will replacce with last row? or option on settings
            #endregion

            try
            {
                #region [Declare and instantiate variables for worksheet/book]
                Excel.Workbook Wkb = xls.ActiveWorkbook;
                Excel.Worksheet Wks;   // get current sheet

                Wks = Wkb.ActiveSheet;

                string strColumnName = CommonExcelClasses.getExcelColumnLetter((int)intColumToCheck);

                DialogResult dlgResult = DialogResult.Yes;

                string strMessage;


                // start of loop
                int intLastRow = CommonExcelClasses.getLastRow(Wks);
                decimal intSourceRow = intStartRow;
                decimal intLastCol = CommonExcelClasses.getLastCol(Wks);
                #endregion

                #region [Display a Message?]
                if (boolDisplayInitialMessage)
                {
                    strMessage = "";
                    strMessage = strMessage + "Worksheet: " + Wks.Name + LF;
                    strMessage = strMessage + "Column: " + strColumnName + LF;
                    strMessage = strMessage + "and: " + strColourOrDelete + " ones which are the same";

                    if (booltimeTaken)
                    {
                        strMessage = strMessage + LF + " and display the time taken";
                    }

                    strMessage = strMessage + "?";

                    dlgResult = MessageBox.Show(strMessage, "Duplicate Rows Check", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    // just for all cols this once
                    // remove formatting - format black and white but only if no was selected
                    if (dlgResult == DialogResult.No)
                        if (boolClearFormatting)
                            CommonExcelClasses.clearFormattingRange(Wks);

                }
                #endregion

                #region [Start of work]
                if (dlgResult == DialogResult.Yes)
                {
                    DateTime dteStart = DateTime.Now;

                    decimal intStartCol = 1;
                    decimal intSourceCol = 1;
                    decimal curNoRecords = 0;

                    decimal[] arrRows = new decimal[100];
                    int intRowArrayPointer = 0;

                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                    #region {start of loop into array}
                    while (!CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow, intColumToCheck]))
                    {
                        // col loop here
                        for (intSourceCol = intStartCol; intSourceCol <= intLastCol; intSourceCol++)
                        {
                            if (Wks.Cells[intSourceRow, intSourceCol].Value != Wks.Cells[intSourceRow + 1, intSourceCol].Value)
                            {
                                break;
                            }
                        }

                        // if all columns were the same
                        if (intSourceCol == (intColumToCheck + intNoCheckCols))
                        {
                            curNoRecords++;

                            if (strColourOrDelete == "Colour")
                            {
                                CommonExcelClasses.colourCells(Wks, (intSourceRow + 1), "Error", intNoCheckCols, clrFoundForeColour, clrFoundBackColour, false);
                                
                            }
                            else if (strColourOrDelete == "Delete")
                            {
                                Wks.Rows[intSourceRow].Delete();
                                intSourceRow--;
                            } else {
                                arrRows[intRowArrayPointer] = (intSourceRow + 1);
                                intRowArrayPointer++;

                                // save row numbers to array to clear after
                                // colourCells(Wks, (intSourceRow + 1), strColourOrDelete, intNoCheckCols, clrFoundForeColour, clrFoundBackColour, false);

                            }

                            if (CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow + 1, intColumToCheck]))
                                break;



                        }

                        intSourceRow++;

                    }
                    #endregion

                    // ok now have an array
                    #region [Deal with result array]
                    if (strColourOrDelete == "Clear")
                    {
                        for (int i = arrRows.GetLowerBound(0); i <= arrRows.GetUpperBound(0); i++)
                        {
                            if (arrRows[i] > 0)
                            {
                                CommonExcelClasses.colourCells(Wks, arrRows[i], strColourOrDelete, intNoCheckCols, clrFoundForeColour, clrFoundBackColour, false);
                            }
                            else
                                break;
                        }


                    }

                    #endregion

                    #region [Say complete?]
                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);

                    if (boolDisplayCompleteMessage)
                    {

                        strMessage = "Complete ...";

                        if (booltimeTaken)
                        {

                            DateTime dteEnd = DateTime.Now;
                            int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                            strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds + LF;
                            strMessage = strMessage + LF + "And handled: " + curNoRecords.ToString() + " duplicates";

                        }
                        CommonExcelClasses.MsgBox(strMessage);
                    }
                    #endregion 

                }
                #endregion

                #region [Release memory]
                Marshal.ReleaseComObject(Wks);
                Marshal.ReleaseComObject(Wkb);
                #endregion

            }
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("There was an error?", "Error");
                Console.WriteLine(excpt.Message);
            }
        }

        // left here - 1gvb9
        internal void compareSheets(Excel.Application xls)
        {
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            string strCompareOrColour = myData.CompareOrColour;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolClearFormatting = myData.ClearFormatting;

            Color clrColourFore_Found = ColorTranslator.FromHtml(myData.ColourFore_Found);
            Color clrColourFore_NotFound = ColorTranslator.FromHtml(myData.ColourFore_NotFound);
            bool boolFontBold_Found = myData.ColourBold_Found;
            bool boolFontBold_NotFound = myData.ColourBold_NotFound;

            Color clrColourBack_Found = ColorTranslator.FromHtml(myData.ColourBack_Found);
            Color clrColourBack_NotFound = ColorTranslator.FromHtml(myData.ColourBack_NotFound);

            int intStartRow = (int)myData.ComparingStartRow;

            bool boolTestCode = myData.TestCode;

            #endregion

            try
            {
                #region [Declare and instantiate variables]
                Excel.Workbook Wkb = xls.ActiveWorkbook;
                Excel.Worksheet Wks1;   // get current sheet
                Excel.Worksheet Wks2;   // get sheet next door

                Wks1 = Wkb.ActiveSheet;
                Wks2 = Wkb.Sheets[Wks1.Index + 1];

                int intSheetLastRow1 = CommonExcelClasses.getLastRow(Wks1);
                int intSheetLastRow2 = CommonExcelClasses.getLastRow(Wks2);
                #endregion

                #region [Declare and instantiate variables for worksheet/book]
                if (intSheetLastRow1 >= intStartRow || intSheetLastRow2 >= intStartRow)
                {

                    #region [Ask to display a Message?]
                    DialogResult dlgResult = DialogResult.Yes;
                    string strMessage;
                                        
                    if (boolDisplayInitialMessage)
                    {
                        strMessage = "Compare: " + Wks1.Name + LF +
                                    " against: " + Wks2.Name + LF +
                                        " and: " + strCompareOrColour + " ones which are the same" + LF +
                                       " (starting at row:" + intStartRow.ToString() + ")";

                        if (booltimeTaken)
                        {
                            strMessage = strMessage + LF + " and display the time taken";
                        }

                        strMessage = strMessage + "?";

                        dlgResult = MessageBox.Show(strMessage, "Compare Sheets", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    }
                  

                    int intLastCol = CommonExcelClasses.getLastCol(Wks1);

                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                    // remove formatting - format black and white but only if no was selected
                    if (dlgResult == DialogResult.No)
                        if (boolClearFormatting)
                            CommonExcelClasses.clearFormattingRange(Wks1);

                    #endregion

                    #region [Start of work]
                    if (dlgResult == DialogResult.Yes)
                    {
                        DateTime dteStart = DateTime.Now;

                        int intTargetRow = 0;
                        int intStartColumToCheck = 1;
                        int intColScore = 0;

                        string strValue1 = "";                        

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
                                for (int intColCount = intStartColumToCheck; intColCount <= intLastCol; intColCount++)
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
                            if (intColScore == intLastCol)
                                CommonExcelClasses.colourCells(Wks1, intSourceRow, strCompareOrColour, intLastCol, clrColourFore_Found, clrColourBack_Found, boolTestCode);
                            else
                                CommonExcelClasses.colourCells(Wks1, intSourceRow, "Error", intLastCol, clrColourFore_NotFound, clrColourBack_NotFound, boolTestCode);

                            intColScore = 0;

                        }

                        if (boolTurnOffScreen)
                            CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);
                        #endregion

                        #region [Display Complete Message]
                        if (boolDisplayCompleteMessage)
                        {
                            strMessage = "";
                            strMessage = strMessage + "Compare Complete ...";

                            if (booltimeTaken)
                            {

                                DateTime dteEnd = DateTime.Now;
                                int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                                strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                            }

                            CommonExcelClasses.MsgBox(strMessage);          // localisation?
                        }
                        #endregion

                    }
                   
                } else {
                    if (boolDisplayCompleteMessage)
                        CommonExcelClasses.MsgBox("No data to compare ...", "Warning");          // localisation?
                }


                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);

                #endregion

                #region [Release memory]
                Marshal.ReleaseComObject(Wks1);
                Marshal.ReleaseComObject(Wks2);
                Marshal.ReleaseComObject(Wkb);
                #endregion

            }
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("Are you on the last sheet?", "Error");
                Console.WriteLine(excpt.Message);
            }
        }


        internal void updateTimeSheet(Excel.Application xls)
        {
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file
            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolSaveLastRowNo = myData.TimeSheetGetRowNo;
            decimal intRowCount = myData.TimeSheetRowNo;
            #endregion

            try
            {
                #region [Declare and instantiate variables for worksheet/book]
                // need to loop entire workbook
                Excel.Workbook Wkb = xls.ActiveWorkbook;
                Excel.Worksheet Wks;

                Wks = Wkb.ActiveSheet;

                // will define start row later
                decimal intLastRow = CommonExcelClasses.getLastRow(Wks);
                int intExaminCol = CommonExcelClasses.getExcelColumnNumber("N");

                string strMessage;
                DialogResult dlgResult = DialogResult.Yes;
                #endregion

                #region [Display a Message?]
                if (boolDisplayInitialMessage)
                {
                    strMessage = "Correct: " + Wks.Name + LF +
                                   " (starting at row:" + intRowCount.ToString() + ")";

                    if (booltimeTaken)
                    {
                        strMessage = strMessage + LF + " and display the time taken";
                    }

                    strMessage = strMessage + "?";

                    dlgResult = MessageBox.Show(strMessage, "Correct Timesheet", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                }
                #endregion
                
                #region [Start of work]
                if (dlgResult == DialogResult.Yes)
                {
                    DateTime dteStart = DateTime.Now;

                    // might put this under an option
                    if (boolTurnOffScreen)
                        CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                    #region [Start of loop]
                    while (intRowCount <= intLastRow)
                    {
                        if (!CommonExcelClasses.isEmptyCell(Wks.Cells[intRowCount, intExaminCol]))
                        {
                            if (boolSaveLastRowNo)
                                startOfWeekCheck(Wks, intRowCount, myData.TimeSheetRowNo);
                            
                            // is it a valid day
                            if (CommonExcelClasses.dayCheck(Wks.Cells[intRowCount, intExaminCol].Value.ToString()))
                            {
                                decimal intStartOfWeekRow = intRowCount;
                                string strDay = Wks.Cells[intRowCount, intExaminCol].Value.ToString();

                                // jump down 1 row
                                intRowCount++;

                                while (Wks.Cells[intRowCount, intExaminCol].Value.ToString() == strDay)
                                {
                                    sortHours(Wks, intRowCount, intStartOfWeekRow);
                                    fixDateCol(Wks, intRowCount, intStartOfWeekRow);
                                    intRowCount++;
                                }

                                intRowCount--;

                                repairTimeRecording(Wks, intStartOfWeekRow + 1, intRowCount);

                                // this will add validation to the entire date section / range
                                CommonExcelClasses.addValidationToColumn(Wks, "Q", intStartOfWeekRow + 1, intRowCount, "=rangeCategory");
                            }
                        }

                        intRowCount++;

                    }

                    #endregion
                    
                    #endregion

                    #region [Display Complete Message]
                    if (boolDisplayCompleteMessage)
                    {
                        strMessage = "";
                        strMessage = strMessage + "Process Complete ...";

                        if (booltimeTaken)
                        {
                            DateTime dteEnd = DateTime.Now;
                            int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                            strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                        }

                        CommonExcelClasses.MsgBox(strMessage);          // localisation?
                    }
                }
                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);
                #endregion
                
                #region [Release memory]
                Marshal.ReleaseComObject(Wks);
                Marshal.ReleaseComObject(Wkb);
                #endregion


            }
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("Ther was an error - around row number" + intRowCount.ToString(), "Error");
                CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);

                Console.WriteLine(excpt.Message);
            }
        }


        public void startOfWeekCheck(Excel.Worksheet Wks, decimal intRowCount, decimal intTimeSheetRowNo)
        {
            if (Wks.Cells[intRowCount, 14].Value.ToString() == "Date")
            {
                if (Wks.Cells[intRowCount, 2].Value.ToString() == "Week")
                {
                    if (intRowCount != intTimeSheetRowNo)
                    {                        
                        myData.TimeSheetRowNo = intRowCount;                // set the value 
                        InformationForSettingsForm.SaveData(myData);
                    }

                }
            }
        }

        // repairs sums at bottom of cols A->M then S:
        public void repairTimeRecording(Excel.Worksheet Wks, decimal intRowStart, decimal intRowEnd)
        {

            decimal intCol_Start = CommonExcelClasses.getExcelColumnNumber("A");
            decimal intCol_End = CommonExcelClasses.getExcelColumnNumber("M");

            // string strDeltaCol, strDeltaRange, strSumString;
            string strDeltaCol;
            string strSumString;
            decimal intDeltaA;

            for (intDeltaA = intCol_Start; intDeltaA <= intCol_End; intDeltaA++)
            {
                strDeltaCol = CommonExcelClasses.getExcelColumnLetter((int)intDeltaA);

                strSumString = CommonExcelClasses.createFormula(strDeltaCol, intRowStart, intRowEnd);
                
                Wks.Cells[intRowEnd + 1, intDeltaA].Value = strSumString;
            }

            // ' finally do S
            intDeltaA = 19;
            strDeltaCol = CommonExcelClasses.getExcelColumnLetter((int)intDeltaA);

            strSumString = CommonExcelClasses.createFormula("S", intRowStart, intRowEnd);
            Wks.Cells[intRowEnd + 1, intDeltaA].Value = strSumString;

        }


        public void fixDateCol(Excel.Worksheet Wks, decimal intRowCount, decimal intStartOfWeekRow)
        {
            // 1st
            // string strDateCol = CommonExcelClasses.getExcelColumnLetter(C_COL_DATE);
            // string strRangeDateCol = "=" + strDateCol + intStartOfWeekRow;
            // Wks.Cells[intRowCount, C_COL_CATEGORY].Value = strRangeDateCol;

            // 2nd
            Wks.Cells[intRowCount, C_COL_DATE].Value = "=" + CommonExcelClasses.getExcelColumnLetter(C_COL_DATE) + intStartOfWeekRow;

        }


        public void sortHours(Excel.Worksheet Wks, decimal intRowCount, decimal intStartOfWeekRow)
        {
            string strSearchCat = "";
            if (!CommonExcelClasses.isEmptyCell(Wks.Cells[intRowCount, C_COL_CATEGORY]))
            {
                strSearchCat = Wks.Cells[intRowCount, C_COL_CATEGORY].Value.Trim().ToString();          // get the category
                Wks.Cells[intRowCount, C_COL_CATEGORY].Value = strSearchCat.Trim();                     // put it back trimmed?
            }

            Excel.Range xlCell;

            if (strSearchCat == "!NON WORKING")
            {
                Wks.Cells[intRowCount, C_COL_TOTAL].Value = "";

            } else {

                // need to put hours sum in if not there
                if (Wks.Cells[intRowCount, C_COL_TOTAL].Value == 0)
                {
                    string strTotalRange = "P" + intRowCount + "-O" + intRowCount;
                    string strTotalRangeSum = "=SUM(" + strTotalRange + ")";

                    Wks.Cells[intRowCount, C_COL_TOTAL].Value = strTotalRangeSum;
                }

                string strRange = "A" + intRowCount.ToString() + ":" + "M" + intRowCount.ToString();
                xlCell = Wks.get_Range(strRange);
                xlCell.ClearContents();

                strSearchCat = transformCat(strSearchCat);

                if (strSearchCat.Length > 0)
                {
                    decimal intTargetCol = searchForValueInHeaderCol(Wks, strSearchCat.Trim(), intStartOfWeekRow);

                    if (intTargetCol > 0)
                    {
                        strRange = "=S" + intRowCount.ToString();
                        Wks.Cells[intRowCount, intTargetCol].Value = strRange;
                    } else {
                        CommonExcelClasses.MsgBox("undefined - check (probably could not find category) ","Error");
                    }

                } 

            }

        }


        public decimal searchForValueInHeaderCol(Excel.Worksheet Wks, string strSearchValue, decimal intWhichScanRow)
        {
            decimal intColNo = 0;

            for (intColNo = 1; intColNo <= 14; intColNo++)
            {
                // check no space
                if (!CommonExcelClasses.isEmptyCell(Wks.Cells[intWhichScanRow, intColNo]))
                    if (Wks.Cells[intWhichScanRow, intColNo].Value.ToUpper() == strSearchValue.ToUpper())
                        break;

            }

            return intColNo;

        }


        public string transformCat(string strSearchCat)
        {
            if (strSearchCat != "!NON WORKING") {

                switch (strSearchCat)
                {
                    case "Treasury":
                        strSearchCat = "GT";
                        break;
                    case "SAP":
                        strSearchCat = "UK Apps";
                        break;
                    case "ProArc":
                        strSearchCat = "UK Apps";
                        break;
                    case "Maximo":
                        strSearchCat = "UK Apps";
                        break;
                    case "SNow":
                        strSearchCat = "System";
                        break;
                    case "Trello":
                        strSearchCat = "General";
                        break;
                    case "Personal":
                        strSearchCat = "OOO";
                        break;
                    case "!Holiday":
                        strSearchCat = "OOO";
                        break;
                }
            }

            return strSearchCat;
        }

        public void deleteBlankLines(Excel.Workbook Wkb, string strMode)
        {
            Excel.Worksheet Wks;
            Wks = Wkb.ActiveSheet;

            string strResultsCol = "F";

            myData = myData.LoadMyData();               // read data from settings file

            bool boolRecordTimes = myData.RecordTimes;
            DateTime dteStart = DateTime.Now;

            if (boolRecordTimes)
            {
                // record times

                Wks.Range[strResultsCol + "1"].Value = dteStart;
            }

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

            } else {
                CommonExcelClasses.MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

            if (boolRecordTimes)
            {
                DateTime dteEnd = DateTime.Now;
                int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                Wks.Range[strResultsCol + "2"].Value = dteEnd ;
                Wks.Range[strResultsCol + "3"].Value = milliSeconds;
            }


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

                string strLastCol = CommonExcelClasses.getExcelColumnLetter(intColLast);

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
