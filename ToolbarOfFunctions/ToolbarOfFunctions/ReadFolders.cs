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

            // read in xml here - grab code from elsewhere
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;

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

                    int gintFileCount = 2;

                    // zap the sheet before we start
                    CommonExcelClasses.zapWorksheet(Wks, 1);

                    // string strPath = cfbd.SelectedPath;
                    directorySearch(cfbd.SelectedPath.ToString(), Wks, gintFileCount, boolExtraDetails, false, strWhichDate);
                    // directorySearch(cfbd.SelectedPath.ToString(), Wks, gintFileCount, boolExtraDetails, true, strWhichDate);

                    // listSubFoldersAndFiles(cfbd.SelectedPath.ToString(), Wks, gintFileCount);

                    writeHeaders(Wks, "FILES", boolExtraDetails, strWhichDate);

                    /*
                    DateTime dteEnd = DateTime.Now;
                    int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;
                    Wks.Range[strWhichColumn + "1"].Value = milliSeconds + " milliSeconds";
                    */

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


        public static void getExtraDetails(string file, Excel.Worksheet Wks, int gintFileCount, string strWhichDate)
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

            Wks.Cells[gintFileCount, 2].value = dteWhichTime;

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


        public static void directorySearch(string root, Excel.Worksheet Wks, int gintFileCount, bool boolExtraDetails, bool isRootItrated, string strWhichDate)
        {

            if (!isRootItrated)
            {
                var rootDirectoryFiles = Directory.GetFiles(root);
                foreach (var file in rootDirectoryFiles)
                {
                    // Console.WriteLine(file);
                    Wks.Cells[gintFileCount, 1].value = file.ToString();

                    if (boolExtraDetails)
                        getExtraDetails(file, Wks, gintFileCount, strWhichDate);

                    gintFileCount++;
                }
            }

            // c# code to stop a macro running - 1GVB1: 15-08-2018

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
                            getExtraDetails(file, Wks, gintFileCount, strWhichDate);

                        gintFileCount++;
                    }
                    directorySearch(directory, Wks, gintFileCount, boolExtraDetails, true, strWhichDate);
                }
            }
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

    }
}
