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

namespace RiggingConsoleApp
{
    class Program
    {
        private static string strFileName = "";

        static void Main(string[] args)
        {
            // The code provided will print ‘Hello World’ to the console.
            // Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

            // CommonExcelClasses.MsgBox("Hello World!", "Information");
            // Console.WriteLine("Hello World!");
            // Console.ReadKey();

            // Go to http://aka.ms/dotnet-get-started-console to continue learning how to build a console app! 

            // FileInfo oFileInfo = new FileInfo(existingFile);
            // string existingFile = "C:\\Work\\Rigging7\\OneSheet";
            // string filePath = existingFile;
            openExcelFile();




        }

        private static void openExcelFile()
        {
            string strFileName = "C:\\Work\\Rigging7\\OneSheet\\FieldsToCells.xlsx";

            FileStream stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader;

            //1. Reading Excel file
            if (Path.GetExtension(strFileName).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            //2. DataSet - The result of each spreadsheet will be created in the result.Tables
            var result = excelReader.AsDataSet();

            /*
            var result = excelReader.AsDataSet(new ExcelDataSetConfiguration() {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() {
                    UseHeaderRow = true } });
                    */


            //3. DataSet - Create column names from first row
            // excelReader.IsFirstRowAsColumnNames = false;

            int rowPosition = 8;
            int columnPosition = 1;

            Console.WriteLine(result.Tables[0].Rows.Count);
            Console.WriteLine(result.Tables[0].Columns.Count);


            string strVal1 = result.Tables[0].Rows[rowPosition][columnPosition].ToString();

            CommonExcelClasses.MsgBox("strVal1: " + strVal1, "Information");


            strVal1 = myCellFormat(strVal1);

            CommonExcelClasses.MsgBox("strVal1: " + strVal1, "Information");


            // qwihtut datasedt
            Console.WriteLine(result.Tables[0].Rows[rowPosition][columnPosition]);

            stream.Close();


        }

        private static string myCellFormat(string strCell)
        {
            strCell = strCell.Replace("\r", " ");
            strCell = strCell.Replace("\n", " ");
            strCell = strCell.Replace("\t", " ");
            strCell = strCell.Replace("  ", " ");
            return strCell;
        }


        /*
            private void processHeaderAndLinesIntoList(string strFileName, List<RiggingHeaderDS> lstRiggingHeaderDS, List<RiggingLinesDS> lstRiggingLinesDS)
            {

                /*

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

        */


        private static void codeeg02()
        {

            FileStream stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader;

            //1. Reading Excel file
            if (Path.GetExtension(strFileName).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            //2. DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();

            //3. DataSet - Create column names from first row
            // excelReader.IsFirstRowAsColumnNames = false;
            int rowPosition = 3;
            int columnPosition = 3;

            DataTable dt = result.Tables[0];
            Console.WriteLine(dt.Rows[rowPosition][columnPosition]);

            // qwihtut datasedt
            Console.WriteLine(result.Tables[0].Rows[rowPosition][columnPosition]);

            Console.WriteLine(result.Tables[0].Rows[rowPosition][columnPosition]);

            // another way
            stream = File.Open(@"C:\Users\Desktop\ExcelDataReader.xlsx", FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReaderNew = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet resultNew = excelReaderNew.AsDataSet();

            DataTable dtNew = result.Tables[0];
            string text = dt.Rows[1][0].ToString();



        }




        private static void codeeg01(){

            string strFileName = "";

            using (var stream = File.Open(strFileName, FileMode.Open, FileAccess.Read))
            {

                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // Choose one of either 1 or 2:
                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                // The result of each spreadsheet is in result.Tables
                }
            }

        }

    }


}