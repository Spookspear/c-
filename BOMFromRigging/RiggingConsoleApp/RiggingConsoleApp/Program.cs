#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.IO;                        // for Directory function
// using ExcelDataReader;
using System.Data;
using RiggingConsoleApp.DAL;
using System.Data.Entity;
using RiggingConsoleApp.DAL.Models;

using OfficeOpenXml;
using System.Runtime.InteropServices;

namespace RiggingConsoleApp
{
    class Program
    {
        public const string GC_DELIVERY_DETAILS = "Delivery Details:";
        public const string GC_ADDITIONAL_ITEMS = "Additional Items (Free Text)";

        static void Main(string[] args)
        {
            StartOfRiggingProcess();

            // testEPPlus();
        }

        
        private static void StartOfRiggingProcess()
        {
            string strPath;

            // strPath = "C:\\Work\\Rigging7\\TwoSheets";
            // strPath = "K:\\Work\\Work\\Rigging7\\Projects\\BOMFromRigging\\Rigging7Workbooks\\OneSheet";
            strPath = "K:\\Work\\Work\\Rigging7\\Projects\\BOMFromRigging\\Rigging7Workbooks\\ExampleSheets";

            // not reading sub folders
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

                // SaveToDb( lstRiggingHeaderDS);

            }

            // save after full population
            SaveToDb(lstRiggingHeaderDS);

            // populateWorksheetFromHeaderAndLinesLists(WksMaster, lstRiggingHeaderDS, intLastRow);

        }

        private static void SaveToDb(List<RiggingHeaderDS> lstRiggingHeaderDS)
        {
            using (var db = new RiggingContext())
            {
                db.RiggingHeaders.AddRange(lstRiggingHeaderDS);
                db.SaveChanges();

                //Database.ExecuteSqlCommand  


            }
        }

        private static void processHeaderAndLinesIntoList(string strFileName, List<RiggingHeaderDS> lstRiggingHeaderDS, List<RiggingLineDS> lstRiggingLinesDS)
        {

            // IExcelDataReader WkbNew;
            // WkbNew = openExcelFile(strFileName);
            // var Wks = WkbNew.AsDataSet().Tables[0];

            // changing this library

            FileInfo fiFilePath = new FileInfo(strFileName);

            ExcelPackage Wkb = new ExcelPackage(fiFilePath);
            ExcelWorksheet Wks = Wkb.Workbook.Worksheets["RR05"];

            int iFr = CommonExcelClasses.searchForValue(Wks, GC_DELIVERY_DETAILS, 1);        // iFr = int Footer Row

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
                    ContactPerson = getExcelValue(Wks, arrAddrHead[0],0),
                    BudgetHolder = getExcelValue(Wks, arrAddrHead[1], 0),
                    VesselLocation = getExcelValue(Wks, arrAddrHead[2], 0),
                    ProjectDepartment = getExcelValue(Wks, arrAddrHead[3], 0),
                    DateRequested = getExcelValue(Wks, arrAddrHead[4], 0),
                    DateRequired = getExcelValue(Wks, arrAddrHead[5], 0),
                    ProjectDuration = getExcelValue(Wks, arrAddrHead[6], 0),
                    SAPCostCode = getExcelValue(Wks, arrAddrHead[7], 0),
                    DeliveryDetails = getExcelValue(Wks, arrAddrFoot[0], 0),
                    Remarks = getExcelValue(Wks, arrAddrFoot[1], 0),
                    ATRWONO = getExcelValue(Wks, arrAddrFoot[2], 0),
                    Vendor = getExcelValue(Wks, arrAddrFoot[3], 0),
                    PONumber = getExcelValue(Wks, arrAddrFoot[4], 0)

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

                    if (!CommonExcelClasses.checkEmptyRange(Wks, strCheckRange, 0))
                    {
                        if (getExcelValue(Wks, arrAddrLines[0],0) == GC_ADDITIONAL_ITEMS)
                        {
                            strLineMainOrAdditional = "A";
                        }
                        else
                        {
                            clsRiggingLines = new RiggingLineDS
                            {
                                HighLevelDesc = getExcelValue(Wks, arrAddrLines[0],0),
                                LowLevelDesc = getExcelValue(Wks, arrAddrLines[1], 0),
                                Quantity = getExcelValue(Wks, arrAddrLines[2], 0),
                                ItemValue = getExcelValue(Wks, arrAddrLines[3], 0),
                                TotalValue = getExcelValue(Wks, arrAddrLines[4], 0),
                                TestProcedure = getExcelValue(Wks, arrAddrLines[5], 0),
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
                // Marshal.FinalReleaseComObject(Wks);
                Wks.Dispose();
                Wkb.Dispose();

                // Marshal.ReleaseComObject(fiFilePath);
                // Marshal.FinalReleaseComObject(WkbNew);
                #endregion


            }
            else
            {
                #region [close the worksheet / workbook]
                // Marshal.FinalReleaseComObject(Wks);
                Wkb.Dispose();

                // Marshal.FinalReleaseComObject(WkbNew);
                #endregion

                CommonExcelClasses.MsgBox("Cant find data");
            }

        }


        private static string getExcelValue(ExcelWorksheet wks, string strAddress, int intOffset)
        {

            // double[] dblArrCoords = CommonExcelClasses.getCoordsFromRange1(strAddress);

            int iCol = strAddress.Col();
            int iRow = strAddress.Row();

            string strRetVal = "";

            if (!CommonExcelClasses.isEmptyCell(wks.Cells[(iRow - intOffset), (iCol - intOffset)]))
            {
                wks.Cells[(iRow - intOffset), (iCol - intOffset)].Value.ToString();

            }
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
            // Marshal.FinalReleaseComObject(Wks);


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


        private static void testEPPlus()
        {

            string strPathFile = "K:\\Work\\Work\\Rigging7\\Projects\\BOMFromRigging\\Rigging7Workbooks\\OneSheet\\FieldsToCells.xlsx";
            FileInfo fiFilePath = new FileInfo(strPathFile);

            using (ExcelPackage Wkb = new ExcelPackage(fiFilePath))
            {
                ExcelWorksheet Wks = Wkb.Workbook.Worksheets["RR05"];


                string strRange = "A5";
                string varContactPersonField = Wks.Cells[strRange.Row(), strRange.Col()].Value.ToString();

                strRange = "A6";
                string varContactPersonValue = Wks.Cells[strRange.Row(), strRange.Col()].Value.ToString();

                CommonExcelClasses.MsgBox("varContactPersonField: " + varContactPersonField);
                CommonExcelClasses.MsgBox("varContactPersonValue: " + varContactPersonValue);


            }

        }



        private static void testEPPlusOld()
        {

            string strPathFile = "K:\\Work\\Work\\Rigging7\\Projects\\BOMFromRigging\\Rigging7Workbooks\\OneSheet\\FieldsToCells.xlsx";


            //Open the workbook (or create it if it doesn't exist)
            // var fi = new FileInfo(@"c:\workbooks\myworkbook.xlsx");
            // var fi = new FileInfo(strPathFile);
            FileInfo fiFilePath = new FileInfo(strPathFile);

            using (ExcelPackage Wkb = new ExcelPackage(fiFilePath))
            {
                //Get the Worksheet created in the previous codesample. 
                // var Wks = Wkb.Workbook.Worksheets["RR05"];
                ExcelWorksheet Wks = Wkb.Workbook.Worksheets["RR05"];


                // Set the cell value using row and column.
                // Wks.Cells[2, 1].Value = "This is cell B1. It is set to bolds";
                string strRange = "";

                strRange = "A5";
                // var varContactPersonField = Wks.Cells[strRange.Row(), strRange.Col()].Value.ToString();

                string varContactPersonField = Wks.Cells[strRange.Row(), strRange.Col()].Value.ToString();

                CommonExcelClasses.MsgBox("varContactPersonField: " + varContactPersonField);

                strRange = "A6";
                string varContactPersonValue = Wks.Cells[strRange.Row(), strRange.Col()].Value.ToString();

                CommonExcelClasses.MsgBox("varContactPersonValue: " + varContactPersonValue);


                //The style object is used to access most cells formatting and styles.
                // Wks.Cells[2, 1].Style.Font.Bold = true;
                //Save and close the package.
                // Wkb.Save();
            }

            /*

            // string path = Server.MapPath("YourFolder/YourFile.xlsx")

            int maxRow = 1048576 //max row in excel
            int maxCol = 0;

            for (int row = 2; row < maxRow; row++)//for looping the rows
            {
                //check if the row has a value by checking the 1st cell of the row if not break the loop
                if (ExcelWorksheet.Cells[row, 1].Value == null)
                    break;
                int temp = 1;
                if (row == 2)
                    for (int col = 1; col <= temp; col++)
                    {
                        if (worksheet.Cells[1, col].Value == null)
                            maxCol = col - 1;
                        else
                            temp++;
                    }
                string var1, var2;
                for (int col = 1; col <= maxCol + 1; col++)//for looping the columns
                {
                    if (col == 1)
                        var1 = worksheet.Cells[row, col].Value.ToString();
                    else //for column 2
                        var2 = worksheet.Cells[row, col].Value.ToString();
                    if (col == maxCol + 1)//condition to insert into database
                                    //here you can call a method that inserts the variables
                                    //ie var 1 and var2 in the database

                }
            }



        private static string Demo()
        {

            string strAddress = "H21";

            int iCol = strAddress.Col();
            int iRow = strAddress.Row();

        }



    */

        }



    }


}