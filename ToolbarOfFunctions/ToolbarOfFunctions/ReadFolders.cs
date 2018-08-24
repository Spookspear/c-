#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using DaveChambers.FolderBrowserDialogEx;
using ToolbarOfFunctions_CommonClasses;
using System.Windows.Forms;
using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo

namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {

        public void readFolders(Excel.Workbook Wkb)
        {
            Excel.Worksheet Wks;            // instantiate worksheet object
            Wks = Wkb.ActiveSheet;          // point to active sheet

            // read in xml here
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolExtractFileName = myData.ExtractFileName;
            decimal intColNoForExtractedFile = myData.ColExtractedFile;

            string strWhichDate = myData.FileDateTime;
            #endregion

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
            // string strWhichColumn = "F";

            if (dlgReadExtraDetails != DialogResult.Cancel)
            {
                if (cfbd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    DateTime dteStart = DateTime.Now;

                    // zap the sheet before we start
                    CommonExcelClasses.zapWorksheet(Wks, 1);

                    fileScan(cfbd.SelectedPath.ToString(), Wks, boolExtraDetails, strWhichDate, boolExtractFileName, intColNoForExtractedFile);

                    writeHeaders(Wks, "FILES", boolExtraDetails, strWhichDate);

                    #region [Display Complete Message]
                    if (boolDisplayCompleteMessage)
                    {
                        string strMessage = "";
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
            }

            // MsgBox("Finished ...");

        }


        public static DialogResult askGetExtraDetails()
        {
            DialogResult dlgResult = MessageBox.Show("Populate extract detail columns?", "Question", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            return dlgResult;
        }


        public static void getExtraDetails(string file, Excel.Worksheet Wks, int intFileNumber, string strWhichDate, bool boolExtractFileName, decimal intColNoForExtractedFile)
        {
            FileInfo oFileInfo = new FileInfo(file);
            FileVersionInfo oFileVersionInfo = FileVersionInfo.GetVersionInfo(file);            // Get file version info LastWriteTime, LastAccessTime

            DateTime dteWhichTime = oFileInfo.LastAccessTime;

            if (strWhichDate == "CreationTime")
                dteWhichTime = oFileInfo.CreationTime;

            if (strWhichDate == "CreationTimeUtc")
                dteWhichTime = oFileInfo.CreationTimeUtc;

            if (strWhichDate == "LastAccessTime")
                dteWhichTime = oFileInfo.LastAccessTime;

            if (strWhichDate == "LastAccessTimeUtc")
                dteWhichTime = oFileInfo.LastAccessTimeUtc;

            if (strWhichDate == "LastWriteTime")
                dteWhichTime = oFileInfo.LastWriteTime;

            if (strWhichDate == "LastWriteTimeUtc")
                dteWhichTime = oFileInfo.LastWriteTimeUtc;

            Wks.Cells[intFileNumber, 2].value = dteWhichTime;

            Wks.Cells[intFileNumber, 3].value = oFileInfo.Length;                               // Size -- .ToString();

            // Wks.Cells[gintFileCount, 4].value = oFileVersionInfo.FileVersion;
            // 0.0.0.0
            // 16.21.0.0

            string strFileVersioninfo = oFileVersionInfo.FileMajorPart.ToString() + "." +
                                                oFileVersionInfo.FileMinorPart.ToString() + "." +
                                                oFileVersionInfo.FileBuildPart.ToString() + "." +
                                                oFileVersionInfo.FilePrivatePart.ToString();


            if (strFileVersioninfo != "0.0.0.0")
                Wks.Cells[intFileNumber, 4].value = strFileVersioninfo;

            if (!boolExtractFileName)
                Wks.Cells[intFileNumber, intColNoForExtractedFile].value = oFileInfo.Name;                                 // File Name Extracted
            else
                Wks.Cells[(intFileNumber), intColNoForExtractedFile].value = extractFileNameOnly(file.ToString());

        }


        private static void fileScan(string strPath, Excel.Worksheet Wks, bool boolExtraDetails, string strWhichDate, bool boolExtractFileName, decimal intColNoForExtractedFile)
        {
            // see if this works first if it does then loop array

            string searchPattern = "*.*";
            string[] arrFiles    = Directory.GetFiles(strPath, searchPattern, SearchOption.AllDirectories);

            for (int i = arrFiles.GetLowerBound(0); i <= arrFiles.GetUpperBound(0); i++)
            {
                // CommonExcelClasses.MsgBox(arrFiles[i]);

                Wks.Cells[(i+2), 1].value = arrFiles[i].ToString();

                if (boolExtraDetails)
                    getExtraDetails(arrFiles[i], Wks, (i+2), strWhichDate, boolExtractFileName, intColNoForExtractedFile);
                else
                {
                    if (boolExtractFileName)
                    {
                        Wks.Cells[(i + 2), intColNoForExtractedFile].value = extractFileNameOnly(arrFiles[i].ToString());
                    }
                }

            }

        }


        private static string extractFileNameOnly(string strFileName)
        {
            string strRetVal = strFileName;
            string strSlash = Convert.ToChar(92).ToString();
            // string strBuild = "";

            // takes in path and returns file name only
            // While InStr(strFileName, "\") > 0
            //     strFileName = Mid(strFileName, (InStr(1, strFileName, "\", vbTextCompare) + 1))
            // Wend

            while (strRetVal.Contains(strSlash) )
            {
                strRetVal = strRetVal.Substring( strRetVal.IndexOf( strSlash ) +1, (strRetVal.Length-strRetVal.IndexOf(strSlash)-1));
            }

            return strRetVal;

        }
    }
}
