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
using System.Linq;

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
        public const string GC_ADDITIONAL_ITEMS = "Additional Items (Free Text)";



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
            if (dlgResult == DialogResult.Yes) {

                DateTime dteStart = DateTime.Now;

                int intLastRow = CommonExcelClasses.getLastRow(WksMaster) + 1;
                int intNoOfFiles = arrFiles.GetUpperBound(0);
                int intObjTotal = (intNoOfFiles+1);

                // delare lists
                List<RiggingHeaderDS> lstRiggingHeaderDS = new List<RiggingHeaderDS>();
                List<RiggingLinesDS> lstRiggingLinesDS = new List<RiggingLinesDS>();

                for (int intFileNo = arrFiles.GetLowerBound(0); intFileNo <= arrFiles.GetUpperBound(0); intFileNo++)
                {
                    processHeaderAndLinesIntoList(arrFiles[intFileNo].ToString(), lstRiggingHeaderDS, lstRiggingLinesDS);

                }

                populateWorksheetFromHeaderAndLinesLists(WksMaster, lstRiggingHeaderDS, intLastRow);
                WksMaster.Columns.AutoFit();

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
            #endregion

        }


        private void processHeaderAndLinesIntoList(string strFileName, List<RiggingHeaderDS> lstRiggingHeaderDS, List<RiggingLinesDS> lstRiggingLinesDS)
        {

            #region [workbook Stuff]
            // open workbook
            var oXL = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false      // change to false on live

            };

            // open workbook
            Workbook WkbToScan = oXL.Workbooks.Open(strFileName.ToString(), ReadOnly: true);

            Excel.Worksheet WksNew = WkbToScan.Worksheets["RR05"];
            int intDAdrRw = CommonExcelClasses.searchForValue(WksNew, GC_DELIVERY_DETAILS, 1);
            int intLineRowStart = 10;

            int intNoLines = ((intDAdrRw - 2) - (intLineRowStart - 1));                     //  Number of lines is addresses of: bottom = ((Delivery Details) - 2 := (28 - 9?) := 19

            string[] arrAddrHead = populateAddressHeader();
            string[] arrAddrFoot = prepareParseAddressArrayFooter(intDAdrRw);               // Read each Footer

            int intLRow = 0;
            #endregion

            #region [Header]

            // add a gui to link the header and lines

            RiggingHeaderDS clsRiggingHeader;
            RiggingLinesDS clsRiggingLines;

            clsRiggingHeader = new RiggingHeaderDS
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

            #endregion

            #region [Line Items]
            lstRiggingLinesDS = new List<RiggingLinesDS>();

            // Line Indicator (Main or Additional) line items
            string strLineMainOrAdditional;
            strLineMainOrAdditional = "H";

            for (int index = 0; index < intNoLines; index++)
            {
                intLRow = (index + intLineRowStart);

                string strCheckRange = "A" + intLRow.ToString() + ":C" + intLRow.ToString();
                string[] arrAddrLines = populateAddressLines(intLRow);

                if (!CommonExcelClasses.checkEmptyRange(WksNew, strCheckRange))
                {
                    if (getExcelValue(WksNew, arrAddrLines[0]) == GC_ADDITIONAL_ITEMS)
                    {
                        strLineMainOrAdditional = "A";
                        // SwitchMainOrAdditional
                        // clsRiggingLines.SwitchMainOrAdditional();
                    }
                    else
                    {
                        clsRiggingLines = new RiggingLinesDS
                        {
                            HighLevelDesc = getExcelValue(WksNew, arrAddrLines[0]),
                            LowLevelDesc = getExcelValue(WksNew, arrAddrLines[1]),
                            Quantity = getExcelValue(WksNew, arrAddrLines[2]),
                            ItemValue = getExcelValue(WksNew, arrAddrLines[3]),
                            TotalValue = getExcelValue(WksNew, arrAddrLines[4]),
                            TestProcedure = getExcelValue(WksNew, arrAddrLines[5]),
                            LineOrAdditional = strLineMainOrAdditional
                        };

                        lstRiggingLinesDS.Add(clsRiggingLines);
                    }

                }

            }
            #endregion

            #region [Assign Lines to Header]
            clsRiggingHeader.lstRiggingLines = lstRiggingLinesDS;
            lstRiggingHeaderDS.Add(clsRiggingHeader);
            #endregion

            #region [close the worksheet / workbook]
            Marshal.FinalReleaseComObject(WksNew);
            WkbToScan.Close(false);
            Marshal.FinalReleaseComObject(WkbToScan);
            #endregion


        }


        private void populateWorksheetFromHeaderAndLinesLists(Worksheet wksMaster, List<RiggingHeaderDS> lstRiggingHeaderDS, int iRow)
        {
            foreach (RiggingHeaderDS h in lstRiggingHeaderDS)
            {
                #region [Line Items]
                // int iCol = 11;

                writeOutRiggingHeader(wksMaster, h, iRow);


                foreach (RiggingLinesDS l in h.lstRiggingLines)
                {
                    // writeOutRiggingHeader(wksMaster, h, iRow);
                    writeOutRiggingLines(wksMaster, l, iRow);

                    iRow++;
                    // iCol = 11;

                }

                // iRow++;
                #endregion

                // iCol = 1;

            }

        }

        private void writeOutRiggingLines(Worksheet wksMaster, RiggingLinesDS l, int iRow)
        {
            int iCol = 11;

            wksMaster.Cells[iRow, iCol].value = l.HighLevelDesc;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.LowLevelDesc;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.Quantity;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.ItemValue;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.TotalValue;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.TestProcedure;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = l.LineOrAdditional;

            iRow++;
            iCol = 11;
        }

        private void writeOutRiggingHeader(Worksheet wksMaster, RiggingHeaderDS h, int iRow)
        {

            int iCol = 1;

            #region [Header]
            wksMaster.Cells[iRow, iCol].value = h.FileName;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.FileDate;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.ContactPerson;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.BudgetHolder;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.VesselLocation;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.ProjectDepartment;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.DateRequested;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.DateRequired;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.ProjectDuration;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.SAPCostCode;
            #endregion

            #region [Footer]
            iCol = 18;
            wksMaster.Cells[iRow, iCol].value = h.DeliveryDetails;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.Remarks;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.ATRWONO;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.Vendor;
            iCol++;
            wksMaster.Cells[iRow, iCol].value = h.PONumber;

            #endregion

        }

        private void populateWorksheetFromHeaderAndLinesListsWorking(Worksheet wksMaster, List<RiggingHeaderDS> lstRiggingHeaderDS, int iRow)
        {
            int iCol = 1;

            foreach (RiggingHeaderDS h in lstRiggingHeaderDS)
            {

                #region [Header]
                wksMaster.Cells[iRow, iCol].value = h.FileName;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.FileDate;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.ContactPerson;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.BudgetHolder;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.VesselLocation;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.ProjectDepartment;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.DateRequested;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.DateRequired;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.ProjectDuration;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.SAPCostCode;
                #endregion

                #region [Footer]
                iCol = 18;
                wksMaster.Cells[iRow, iCol].value = h.DeliveryDetails;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.Remarks;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.ATRWONO;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.Vendor;
                iCol++;
                wksMaster.Cells[iRow, iCol].value = h.PONumber;

                #endregion

                #region [Line Items]
                iCol = 11;

                foreach (RiggingLinesDS l in h.lstRiggingLines)
                {
                    wksMaster.Cells[iRow, iCol].value = l.HighLevelDesc;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.LowLevelDesc;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.Quantity;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.ItemValue;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.TotalValue;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.TestProcedure;
                    iCol++;
                    wksMaster.Cells[iRow, iCol].value = l.LineOrAdditional;

                    iRow++;
                    iCol = 11;

                }

                // iRow++;
                #endregion

                iCol = 1;

            }

        }


        private string getExcelValue(Excel.Worksheet WksNew,  string v)
        {

            Excel.Range xlCell;

            xlCell = WksNew.get_Range(v);

            string strRetVal = "";

            if (!CommonExcelClasses.isEmptyCell(xlCell, false))
                strRetVal = xlCell.Value.ToString();

            return strRetVal;


        }



        private string[] populateAddressLines(int intLRow)
        {

            string[] arrAddrLines = { "A" + intLRow.ToString(),         // Ax   High Level Description      A10
                                      "B" + intLRow.ToString() ,        // Bx   Low Level Description       B10
                                      "C" + intLRow.ToString() ,        // Cx   Quantity                    C10
                                      "D" + intLRow.ToString() ,        // Dx   Item Value                  D10
                                      "E" + intLRow.ToString(),         // Ex   Total Value                 E10
                                      "F" + intLRow.ToString(),         // Fx   Test Procedure/Legislation  E10
                                    };


            return arrAddrLines;


        }


        private string[] populateAddressHeader()
        {
            string[] arrAddresses = { "A6", "B6", "C6", "E6", "A8", "B8", "C8", "E8" };     // will eventually read this from somewhere
            return arrAddresses;

        }



        private string[] prepareParseAddressArrayFooter(int intDAdrRw)
        {

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
