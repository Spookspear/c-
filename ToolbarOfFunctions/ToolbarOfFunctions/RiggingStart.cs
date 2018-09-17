#pragma warning disable IDE1006 // Naming Styles

using System;
// using System.Collections.Generic;
// using System.Linq;
// using System.Text;
// using System.Threading.Tasks;
using System.IO;                        // for Directory function
using System.Windows.Forms;             // for ok prompt
using System.Diagnostics;               // .FileVersionInfo

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;


using ToolbarOfFunctions_CommonClasses;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using System.Collections.Generic;

namespace ToolbarOfFunctions
{
    /// <summary>
    /// 11-09-2018
    ///  This will evetually be moved into a console applicaiton 
    ///  to be run as a sevice
    ///  Things outstanding
    ///     write data to a SQL database
    ///     Read documents from sharepoint / Documentset
    /// </summary>
    /// 


    public partial class ThisAddIn
    {
        // Start of rigging
        // need to read a folder of excel files
        // into a database
        // but only once
        // and if updated

        public const string GC_DELIVERY_DETAILS = "Delivery Details:";


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

            string strPath = "C:\\Work\\Rigging7\\OneSheet";

            // read into a worksheet initally
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet WksMaster = Wkb.ActiveSheet;                    // get the sheet we are on  // point to active sheet

            // check not reading sub folders
            string searchPattern = "*.xlsx";
            string[] arrFiles = Directory.GetFiles(strPath, searchPattern, SearchOption.TopDirectoryOnly);
            #endregion

            #region [Ask to display a Message?]
            // DialogResult dlgResult = DialogResult.Yes;
            // string strMessage;

            /*
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
            */


            #endregion

            #region [Start of work]
            // if (dlgResult == DialogResult.Yes) {

                DateTime dteStart = DateTime.Now;
                // open up each sheet - 1st
                // will move to own sub later
                // put into this sheet

                int intLastRow = CommonExcelClasses.getLastRow(WksMaster);
                int intNoOfFiles = arrFiles.GetUpperBound(0);
                int intObjTotal = (intNoOfFiles+1);
                int intObjCnt = 1;

                // need to turn riggingDS into array as well
                // guessing obj cantr be zero based and always has to be plus 1
                RiggingHeaderDS[] riggingDS = new RiggingHeaderDS[10];
                // RiggingHeaderDS.RiggingLinesDS[] linesDs = new RiggingHeaderDS.RiggingLinesDS[300];

                for (int intFileNo = arrFiles.GetLowerBound(0); intFileNo <= arrFiles.GetUpperBound(0); intFileNo++)
                {
                    processHeaderAndLineItemsIntoObject(arrFiles[intFileNo].ToString(), intFileNo, riggingDS);

                }

                // processObjectIntoWorksheet(WksMaster, intLastRow, intLastRow, riggingDS[intNoOfFiles]);
                // WksMaster.Columns.AutoFit();
            #endregion
            /*
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

                */
            // }

        }


        private void processHeaderAndLineItemsIntoObject(string strFileName, int intFileNo, RiggingHeaderDS[] riggingDS)
        {

            // open workbook
            var oXL = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true      // change to false on live

            };

            // open workbook
            Workbook WkbToScan = oXL.Workbooks.Open(strFileName.ToString(), ReadOnly: true);

            Excel.Worksheet WksNew = WkbToScan.Worksheets["RR05"];
            int intDAdrRw = CommonExcelClasses.searchForValue(WksNew, GC_DELIVERY_DETAILS, 1);
            int intLineStart = 10;
            //  Number of lines is addresses of: bottom = ((Delivery Details) - 2 := (28 - 9?) := 19
            int intNoLines = ((intDAdrRw - 2) - (intLineStart - 1));

            string[] arrAddrHead = populateAddressHeader();
            string[] arrAddrFoot = prepareParseAddressArrayFooter(intDAdrRw);            // Read each Footer

            string strRange = "";
            int intLRow;


            riggingDS[intFileNo] = new RiggingHeaderDS
            {
                FileName = strFileName,
                FileDate = CommonExcelClasses.getFileDate(strFileName.ToString()),
                ContactPerson = getExcelValue(WksNew, arrAddrHead[0]),
                BudgetHolder = getExcelValue(WksNew, arrAddrHead[1]),
                VesselLocation = getExcelValue(WksNew, arrAddrHead[2]),
                ProjectDepartment = getExcelValue(WksNew, arrAddrHead[3]),
                DateRequested = getExcelValue(WksNew, arrAddrHead[4]),
                DateRequired = getExcelValue(WksNew, arrAddrHead[5]),
                ProjectDuration = getExcelValue(WksNew, arrAddrHead[6]),
                SAPCostCode = getExcelValue(WksNew, arrAddrHead[7]),
                DeliveryDetails = getExcelValue(WksNew, arrAddrFoot[0]),
                Remarks = getExcelValue(WksNew, arrAddrFoot[1]),
                ATRWONO = getExcelValue(WksNew, arrAddrFoot[2]),
                Vendor = getExcelValue(WksNew, arrAddrFoot[3]),
                PONumber = getExcelValue(WksNew, arrAddrFoot[4])

            };


            // RiggingHeaderDS.RiggingLinesDS[] linesDs = new RiggingHeaderDS.RiggingLinesDS[intNoLines];

            // riggingDS.RiggingLinesDS[] = testc new riggingDS.RiggingLinesDS[];

            // riggingDS[intFileNo].RiggingLinesDS riggingLinesDs = new riggingDS[].Length;

            // Student student = new Student();
            // student.Grade.GradeName = "A";

            // Grade Grade = new Grade();
            // Grade.Student.Name = "test";

            // student.Grade.GradeName = "A";

            // linesDs.RiggingLinesDS.Quantity = "1";
            // RiggingHeaderDS[] riggingDS = new RiggingHeaderDS[10];
            // riggingDS[intFileNo].RiggingLinesDS.Quantity = "1";
            // RiggingHeaderDS.RiggingLinesDS[] linesDs = new RiggingHeaderDS.RiggingLinesDS[300];
            // RiggingHeaderDS.[] linesDs = new RiggingHeaderDS[10];
            // RiggingHeaderDS linesDs = new RiggingHeaderDS();
            // RiggingLinesDS[] linesDSTst = riggingDS.RiggingLinesDS;


            RiggingLinesDS[] ds = new RiggingLinesDS[30];

            // ds[0].HighLevelDesc = "";

            for (int index = 0; index <= intNoLines; index++)
            {

                // instantiate obj
                // linesDs[index] = new riggingDS[intFileNo].RiggingLinesDS();

                intLRow = (index + intLineStart);
                strRange = "A" + intLRow.ToString();

                ds[index] = new RiggingLinesDS();

                // riggingDS[intFileNo].RiggingLinesDS.HighLevelDesc = getExcelValue(WksNew, strRange);
                ds[index].HighLevelDesc = getExcelValue(WksNew, strRange);

                strRange = "B" + intLRow.ToString();
                // riggingDS[intFileNo].RiggingLinesDS.LowLevelDesc = getExcelValue(WksNew, strRange);
                ds[index].LowLevelDesc = getExcelValue(WksNew, strRange);

                strRange = "C" + intLRow.ToString();
                // riggingDS[intFileNo].RiggingLinesDS.Quantity = getExcelValue(WksNew, strRange);
                ds[index].Quantity = getExcelValue(WksNew, strRange);

                strRange = "D" + intLRow.ToString();
                // riggingDS[intFileNo].RiggingLinesDS.ItemValue = getExcelValue(WksNew, strRange);
                ds[index].ItemValue = getExcelValue(WksNew, strRange);

                // ds[index].RiggingHeaderDS.Add(riggingDS[intFileNo]);


                // riggingDS[intFileNo].RiggingLinesDS.GetTotalValue();


            }
        


            // close workbook
            Marshal.FinalReleaseComObject(WksNew);
            WkbToScan.Close(false);
            Marshal.FinalReleaseComObject(WkbToScan);


        }








        private void processLineItemsIntoObject(Workbook wkbToScan, int intFileNo, RiggingHeaderDS[] riggingDS)
        {
            // starting at: A10 
            // loop down
            // creating new objects for each line
            Excel.Worksheet WksNew = wkbToScan.Worksheets["RR05"];

            // find: Additional Items (Free Text)
            // will need number of lines
            int intDAdrRw = CommonExcelClasses.searchForValue(WksNew, GC_DELIVERY_DETAILS, 1);

            int intLineStart = 10;

            //  Number of lines is addresses of: bottom = ((Delivery Details) - 2 := (28 - 9?) := 19
            int intNoLines = ((intDAdrRw - 2) - (intLineStart-1));

            // riggingDS[0].

            // working
            // RiggingHeaderDS.RiggingLinesDS[] linesDs = new RiggingHeaderDS.RiggingLinesDS[intNoLines];
            // RiggingHeaderDS.RiggingLinesDS[] linesDs = new RiggingHeaderDS.RiggingLinesDS[intNoLines];

            // how do I link these to the header?
            // do I need to?


        }

        private void processObjectIntoWorksheet(Excel.Worksheet wksMaster, int iRow, int iOffSet, RiggingHeaderDS riggingDS)
        {
            int iR = iRow + iOffSet;
            int iC = 1;

            // header
            wksMaster.Cells[iR , iC].value = riggingDS.FileName;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.FileDate;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.ContactPerson;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.BudgetHolder;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.VesselLocation;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.ProjectDepartment;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.DateRequested;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.DateRequired;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.ProjectDuration;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.SAPCostCode;


            // the line items go here // along with the exisiting data
            // the extrac items - wont matter on a database though


            iC = 18;

            // footer
            wksMaster.Cells[iR , iC].value = riggingDS.DeliveryDetails;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.Remarks;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.ATRWONO;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.Vendor;
            iC++;
            wksMaster.Cells[iR , iC].value = riggingDS.PONumber;


        }

        private string getExcelValue(Excel.Worksheet WksNew,  string v)
        {

            Excel.Range xlCell;

            xlCell = WksNew.get_Range(v);

            string strRetVal = "";

            if (!CommonExcelClasses.isEmptyCell(xlCell))
                strRetVal = xlCell.Value.ToString();

            return strRetVal;


        }

        private string[] populateAddressHeader()
        {
            string[] arrAddresses = { "A6", "B6", "C6", "E6", "A8", "B8", "C8", "E8" };     // will eventually read this from somewhere
            return arrAddresses;

        }



        private string[] prepareParseAddressArrayFooter(int intDAdrRw)
        {

            // Excel.Worksheet WksNew = wkbToScan.Worksheets["RR05"];
            // int intDAdrRw = CommonExcelClasses.searchForValue(WksNew, GC_DELIVERY_DETAILS, 1);

            string[] arrFooterAddr = { "B" + intDAdrRw.ToString() ,           // Bx           Delivery Details       B30
                                       "B" + (intDAdrRw +2).ToString() ,      // Bx+2         Remarks                B32
                                       "A" + (intDAdrRw +5).ToString() ,      // Ax+5         ATR WO NO              A35
                                       "B" + (intDAdrRw +5).ToString() ,      // Bx+5         Vendor                 B35
                                       "D" + (intDAdrRw +5).ToString()  };    // Dx+5         PO Number              D35


            return arrFooterAddr;

            // close and free the memory
            // Marshal.FinalReleaseComObject(WksNew);


        }


        private void readRiggingLinesIntoWorkbook(Excel.Worksheet wksMaster, Workbook wkbToScan, int intRowCount, string[] arrAddresses)
        {


        }



    }
}
