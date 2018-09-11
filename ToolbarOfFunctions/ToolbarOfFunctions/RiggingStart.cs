﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;            // for Directory function
using System.Windows.Forms;                 // for ok prompt

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;


using ToolbarOfFunctions_CommonClasses;
using System.Runtime.InteropServices;

namespace ToolbarOfFunctions
{
    // class RiggingStart
    // {
    public partial class ThisAddIn
    {
        // Start of rigging
        // need to read a folder of excel files
        // into a database
        // but only once
        // and if updated

        /// <summary>
        /// 11-09-2018
        ///  This will evetually be moved into a console applicaiton 
        ///  to be run as a sevice
        ///  Things outstanding
        ///     write data to a SQL database
        ///     Read documents from sharepoint / Documentset
        /// </summary>

        public void StartOfRiggingProcess(Excel.Application xls)
        {

            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();
            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            #endregion

            #region [Declare and instantiate work book/sheet variables]

            string strPath = "C:\\Work\\Rigging7\\ExampleSheets";

            // read into a worksheet initally
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet WksMaster = Wkb.ActiveSheet;            // get the sheet we are on  // point to active sheet

            // check not reading sub folders
            string searchPattern = "*.xlsx";
            string[] arrFiles = Directory.GetFiles(strPath, searchPattern, SearchOption.TopDirectoryOnly);

            // pass this into its own procedure, getting the line number back?
            string[] arrAddresses = { "A6", "B6", "C6", "E6", "A8", "B8", "C8", "E8" };     // will eventually read this from somewhere

            #endregion

            #region [Ask to display a Message?]
            DialogResult dlgResult = DialogResult.Yes;
            string strMessage;

            if (boolDisplayInitialMessage)
            {
                strMessage = "Read workbooks - located : " + strPath + LF +
                            " into: " + WksMaster.Name + LF;

                if (booltimeTaken)
                {
                    strMessage = strMessage + LF + " and display the time taken";
                }

                strMessage = strMessage + "?";

                dlgResult = MessageBox.Show(strMessage, "Read Rigging", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            #endregion

            #region [Start of work]
            if (dlgResult == DialogResult.Yes)
            {
                DateTime dteStart = DateTime.Now;
                // open up each sheet - 1st
                // will move to own sub later
                // put into this sheet
                for (int intRowCount = arrFiles.GetLowerBound(0); intRowCount <= arrFiles.GetUpperBound(0); intRowCount++)
                {
                    // CommonExcelClasses.MsgBox(arrFiles[intRowCount]);

                    // open workbook
                    var oXL = new Microsoft.Office.Interop.Excel.Application
                    {
                        Visible = false      // change to false on live
                    };

                    WksMaster.Cells[(intRowCount + 3), 1].value = arrFiles[intRowCount].ToString();

                    Workbook WkbNew = oXL.Workbooks.Open(arrFiles[intRowCount].ToString(), ReadOnly: true);

                    ReadRiggingHeaderIntoWorkbook(WksMaster, WkbNew, intRowCount, arrAddresses);        // Read each workbook
                    ReadRiggingLinesIntoWorkbook(WksMaster, WkbNew, intRowCount, arrAddresses);         // Read each line


                    WkbNew.Close(false);
                    Marshal.FinalReleaseComObject(WkbNew);

                }
                WksMaster.Columns.AutoFit();
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

                    CommonExcelClasses.MsgBox(strMessage); 
                }
                #endregion

            }

            
        }

        private void ReadRiggingLinesIntoWorkbook(Excel.Worksheet wksMaster, Workbook wkbNew, int intRowCount, string[] arrAddresses)
        {


        }

        private void ReadRiggingHeaderIntoWorkbook(Excel.Worksheet wksMaster, Workbook wkbNew, int intRowCount, string[] arrAddresses)
        {

            // instantiate the needed sheet
            Excel.Worksheet WksNew = wkbNew.Worksheets["RR05"];

            // read cells into worksheet / read in relevant data 

            string strAddress = "";
            for (int intaddresses = arrAddresses.GetLowerBound(0); intaddresses <= arrAddresses.GetUpperBound(0); intaddresses++)
            {
                // need to handle null cells
                strAddress = arrAddresses[intaddresses];

                // create a routine that handles nulls and passes back values

                if (!CommonExcelClasses.isEmptyCell(WksNew.get_Range(strAddress)))
                    wksMaster.Cells[(intRowCount + 3), (3 + intaddresses)].value = WksNew.get_Range(strAddress);

            }

            // close and free the memory
            Marshal.FinalReleaseComObject(WksNew);





        }

        /*
        private static Var GetExcelValue(Excel.Worksheet wksNew, string v)
        {
            Excel.Range xlCell;
            if ( !CommonExcelClasses.isEmptyCell( wksNew.get_Range(v) ) )
                xlCell = wksNew.get_Range(v);          
            return xlCell.ToString();
        } */

        private static void FileScanIntoExcel(string strPath, Excel.Worksheet Wks, bool boolExtraDetails, string strWhichDate, bool boolExtractFileName, decimal intColNoForExtractedFile)
        {
            // see if this works first if it does then loop array
            string searchPattern = "*.*";
            string[] arrFiles = Directory.GetFiles(strPath, searchPattern, SearchOption.AllDirectories);

            for (int i = arrFiles.GetLowerBound(0); i <= arrFiles.GetUpperBound(0); i++)
            {
                CommonExcelClasses.MsgBox(arrFiles[i]);

                Wks.Cells[(i + 2), 1].value = arrFiles[i].ToString();

                if (boolExtraDetails)
                    getExtraDetails(arrFiles[i], Wks, (i + 2), strWhichDate, boolExtractFileName, intColNoForExtractedFile);
                else
                {
                    if (boolExtractFileName)
                    {
                        Wks.Cells[(i + 2), intColNoForExtractedFile].value = extractFileNameOnly(arrFiles[i].ToString());
                    }
                }

            }

        }



    }
}
