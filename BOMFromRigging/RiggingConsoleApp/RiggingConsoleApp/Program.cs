#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.IO;                        // for Directory function
// using System.Windows.Forms;             // for ok prompt
using System.Diagnostics;               // .FileVersionInfo

// using Excel = Microsoft.Office.Interop.Excel;
// using Microsoft.Office.Interop.Excel;


// using RiggingConsoleApp_CommonExcelClasses;


using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelDataReader;
using System.Data;
using RiggingConsoleApp.DAL;

using System.Data.Entity;

using RiggingConsoleApp.DAL.Models;



namespace RiggingConsoleApp
{
    class Program
    {
        private static string strFileName = "";


        public const int _COL1 = 0;     // A
        public const int _ROW1 = 1;     // 6

        public const int _COL2 = 2;
        public const int _ROW2 = 3;

        public const string GC_DELIVERY_DETAILS = "Delivery Details:";
        public const string GC_ADDITIONAL_ITEMS = "Additional Items (Free Text)";


        static void Main(string[] args)
        {

            StartOfRiggingProcess();

        }

        private static void StartOfRiggingProcess()
        {
            // string strMsg = "hello Nicola".Message();

            string strPath = "C:\\Work\\Rigging7\\TwoSheets";


            
            // check not reading sub folders
            string strSearchPattern = "*.xlsx";
            string[] arrOfFiles = Directory.GetFiles(strPath, strSearchPattern, SearchOption.TopDirectoryOnly);

            DateTime dteStart = DateTime.Now;           // for time recording later

            int intNoOfFiles = arrOfFiles.GetUpperBound(0);

            // declare lists
            List<RiggingHeaderDS> lstRiggingHeaderDS = new List<RiggingHeaderDS>();
            List<RiggingLineDS> lstRiggingLinesDS = new List<RiggingLineDS>();

            for (int intFileNo = arrOfFiles.GetLowerBound(0); intFileNo <= arrOfFiles.GetUpperBound(0); intFileNo++)
            {
                processHeaderAndLinesIntoList(arrOfFiles[intFileNo].ToString(), lstRiggingHeaderDS, lstRiggingLinesDS);

                SaveToDb( lstRiggingHeaderDS);

            }

            // populateWorksheetFromHeaderAndLinesLists(WksMaster, lstRiggingHeaderDS, intLastRow);

        }

        private static void SaveToDb(List<RiggingHeaderDS> lstRiggingHeaderDS)
        {
            using (var db = new RiggingContext())
            {

                db.RiggingHeaders.AddRange(lstRiggingHeaderDS);
                db.SaveChanges();
   
            }
        }

        private static void processHeaderAndLinesIntoList(string strFileName, List<RiggingHeaderDS> lstRiggingHeaderDS, List<RiggingLineDS> lstRiggingLinesDS)
        {

            IExcelDataReader WkbNew;

            WkbNew = openExcelFile(strFileName);
            var WksNew = WkbNew.AsDataSet().Tables[0];

            // var WksNew = varWksNew.Tables[0];

            int iFr = (CommonExcelClasses.searchForValue(WksNew, GC_DELIVERY_DETAILS, 0) +1);       // iFr = int Footer Row

            if (iFr > 1)
            {
                int intLineRowStart = 10;

                int intNoLines = ((iFr - 2) - (intLineRowStart - 1));                     //  Number of lines is addresses of: bottom = ((Delivery Details) - 2 := (28 - 9?) := 19

                string[] arrAddrHead = populateAddressHeader();
                string[] arrAddrFoot = prepareParseAddressArrayFooter(iFr);               // Read each Footer

                int intLRow = 0;

                #region [Header]

                RiggingHeaderDS clsRiggingHeader;
                RiggingLineDS clsRiggingLines;

                clsRiggingHeader = new RiggingHeaderDS
                {
                    FileName = strFileName,
                    FileDate = CommonExcelClasses.getFileDate(strFileName.ToString()),
                    ContactPerson = getExcelValue(WksNew, arrAddrHead[0],1),
                    BudgetHolder = getExcelValue(WksNew, arrAddrHead[1], 1),
                    VesselLocation = getExcelValue(WksNew, arrAddrHead[2], 1),
                    ProjectDepartment = getExcelValue(WksNew, arrAddrHead[3], 1),
                    DateRequested = getExcelValue(WksNew, arrAddrHead[4], 1),
                    DateRequired = getExcelValue(WksNew, arrAddrHead[5], 1),
                    ProjectDuration = getExcelValue(WksNew, arrAddrHead[6], 1),
                    SAPCostCode = getExcelValue(WksNew, arrAddrHead[7], 1),
                    DeliveryDetails = getExcelValue(WksNew, arrAddrFoot[0], 1),
                    Remarks = getExcelValue(WksNew, arrAddrFoot[1], 1),
                    ATRWONO = getExcelValue(WksNew, arrAddrFoot[2], 1),
                    Vendor = getExcelValue(WksNew, arrAddrFoot[3], 1),
                    PONumber = getExcelValue(WksNew, arrAddrFoot[4], 1)

                };

                #endregion

                #region [Line Items]
                lstRiggingLinesDS = new List<RiggingLineDS>();

                // Line Indicator (Main or Additional) line items
                string strLineMainOrAdditional;
                strLineMainOrAdditional = "H";

                for (int index = 0; index < intNoLines; index++)
                {
                    intLRow = (index + intLineRowStart);

                    string strCheckRange = "A" + intLRow.ToString() + ":F" + intLRow.ToString();
                    string[] arrAddrLines = populateAddressLines(intLRow);

                    // this isnt working the way it sould be

                    if (!CommonExcelClasses.checkEmptyRange(WksNew, strCheckRange,1))
                    {
                        if (getExcelValue(WksNew, arrAddrLines[0],1) == GC_ADDITIONAL_ITEMS)
                        {
                            strLineMainOrAdditional = "A";
                        }
                        else
                        {
                            clsRiggingLines = new RiggingLineDS
                            {
                                HighLevelDesc = getExcelValue(WksNew, arrAddrLines[0],1),
                                LowLevelDesc = getExcelValue(WksNew, arrAddrLines[1],1),
                                Quantity = getExcelValue(WksNew, arrAddrLines[2],1),
                                ItemValue = getExcelValue(WksNew, arrAddrLines[3],1),
                                TotalValue = getExcelValue(WksNew, arrAddrLines[4],1),
                                TestProcedure = getExcelValue(WksNew, arrAddrLines[5],1),
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
                // Marshal.FinalReleaseComObject(WksNew);
                WkbNew.Close();

                // Marshal.FinalReleaseComObject(WkbNew);
                #endregion


            }
            else
            {
                #region [close the worksheet / workbook]
                // Marshal.FinalReleaseComObject(WksNew);
                WkbNew.Close();

                // Marshal.FinalReleaseComObject(WkbNew);
                #endregion

                CommonExcelClasses.MsgBox("Cant find data");
            }

        }

        private static string getExcelValue(DataTable dataTable, string strAddress, int intOffset)
        {

            // double[] dblArrCoords = CommonExcelClasses.getCoordsFromRange1(strAddress);

            int iCol = strAddress.Col();
            int iRow = strAddress.Row();


            string strRetVal = dataTable.Rows[(iRow - intOffset)][(iCol - intOffset)].ToString();

            return strRetVal;


        }


        private static string[] populateAddressHeader()
        {
            string[] arrAddresses = { "A6", "B6", "C6", "E6", "A8", "B8", "C8", "E8" };     // will eventually read this from somewhere
            return arrAddresses;

        }

        private static string[] prepareParseAddressArrayFooter(int iFr)
        {
            // iFr = int Footer Row

            string[] arrFooterAddr = { "B" + iFr.ToString() ,           // Bx           Delivery Details       B30
                                       "B" + (iFr +2).ToString() ,      // Bx+2         Remarks                B32
                                       "A" + (iFr +5).ToString() ,      // Ax+5         ATR WO NO              A35
                                       "B" + (iFr +5).ToString() ,      // Bx+5         Vendor                 B35
                                       "D" + (iFr +5).ToString()  };    // Dx+5         PO Number              D35


            return arrFooterAddr;

            // close and free the memory
            // Marshal.FinalReleaseComObject(WksNew);


        }

        private static string[] populateAddressLines(int iLr)
        {
            // iLr = int List Row

            string[] arrAddrLines = { "A" + iLr.ToString(),         // Ax   High Level Description      A10
                                      "B" + iLr.ToString() ,        // Bx   Low Level Description       B10
                                      "C" + iLr.ToString() ,        // Cx   Quantity                    C10
                                      "D" + iLr.ToString() ,        // Dx   Item Value                  D10
                                      "E" + iLr.ToString(),         // Ex   Total Value                 E10      
                                      "F" + iLr.ToString(),         // Fx   Test Procedure/Legislation  E10      
                                    };


            return arrAddrLines;


        }

        private static IExcelDataReader openExcelFile(string strFileName)
        {
            // string strFileName = "C:\\Work\\Rigging7\\OneSheet\\FieldsToCells.xlsx";

            FileStream stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader wkbNew;

            //1. Reading Excel file
            if (Path.GetExtension(strFileName).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                wkbNew = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                wkbNew = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            //2. DataSet - The result of each spreadsheet will be created in the result.Tables
            // var result = excelReader.AsDataSet();

            // return excelReader.AsDataSet();
            return wkbNew;


        }


    }


}