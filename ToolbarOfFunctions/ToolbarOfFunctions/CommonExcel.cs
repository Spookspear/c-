﻿#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

// using Microsoft.Office.Tools.Excel;
// using Microsoft.Office.Interop.Excel;

using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using DaveChambers.FolderBrowserDialogEx;

using System.ComponentModel;
using System.Data;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using ToolbarOfFunctions;

using Microsoft.Office.Core;

using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using System.Data.SqlTypes;



namespace ToolbarOfFunctions_CommonClasses
{
    public class CommonExcelClasses
    {

        public static void ButtonUpdateLabel(RibbonButton rbnButton, string strText)
        {
            rbnButton.Label = strText;
        }


        public static void SplitButtonUpdateLabel(RibbonSplitButton rbnSplitButton, string strText)
        {
            rbnSplitButton.Label = strText;
        }        


        public static void ButtonSetSize(RibbonButton rbnButton, bool boolLargeButton)
        {
            if (boolLargeButton)
            {
                rbnButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            }
            else
            {
                rbnButton.ControlSize = RibbonControlSize.RibbonControlSizeRegular;
            }
        }


        public static bool IsDate(object Expression)
        {
            if (Expression != null)
            {
                if (Expression is DateTime)
                {
                    return true;
                }
                if (Expression is string)
                {
                    // DateTime time1;
                    // return DateTime.TryParse((string)Expression, out time1);
                    return DateTime.TryParse((string)Expression, out DateTime time1);
                }
            }
            return false;
        }

        public static void SplitButtonSetSize(RibbonSplitButton rbnSplitButton, bool boolLargeButton)
        {
            if (boolLargeButton)
                rbnSplitButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            else
                rbnSplitButton.ControlSize = RibbonControlSize.RibbonControlSizeRegular;
        }

        

        public static void MsgBox(string strMessage, string strWhichIcon = "Information")
        {
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

                case "Warning":
                    whichIcon = MessageBoxIcon.Warning;
                    break;

            }

            MessageBox.Show(strMessage, strCaption, MessageBoxButtons.OK, whichIcon);
        }


        public static bool isEmptyCell(Excel.Range xlCell)
        {
            bool boolRetVal = false;

            if (xlCell == null || xlCell.Value2 == null || xlCell.Value2.ToString() == "")
            {
                boolRetVal = true;
            }

            return boolRetVal;

        }

        private static bool WorksheetExist(Excel.Workbook Wkb, string strSheetName)
        {
            bool found = false;

            foreach (Excel.Worksheet Wks in Wkb.Sheets)
            {

                MsgBox("inside WorksheetExist() - ws Name " + Wks.Name.ToLower());

                if (Wks.Name.ToLower() == strSheetName.ToLower())
                {
                    found = true;
                    break;
                }
            }

            return found;
        }

        public static int searchForValue(Excel.Worksheet Wks2, string searchString, int intStartColumToCheck)
        {
            int intRetVal;

            Excel.Range colRange = Wks2.Columns["A:A"];             //get the range object where you want to search from

            Excel.Range resultRange = colRange.Find(

                    What: searchString,

                    LookIn: Excel.XlFindLookIn.xlValues,

                    LookAt: Excel.XlLookAt.xlPart,

                    // SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchOrder: Excel.XlSearchOrder.xlByColumns,

                    SearchDirection: Excel.XlSearchDirection.xlNext

                );                                                  // search searchString in the range, if find result, return a range

            if (resultRange is null)
            {
                intRetVal = 0;
            }
            else
            {
                intRetVal = resultRange.Row;
            }

            return intRetVal;

        }

        public static void deleteEmptyRows(Excel.Worksheet Wks)
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

        public static string getExcelColumnLetter(int columnNumber)
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


        public static int getExcelColumnNumber(string strLetter)
        {
            strLetter = strLetter.ToUpper();
            int intOutNum = 0;

            for (int i = 0; i < strLetter.Length; i++)
            {
                intOutNum *= 26;
                intOutNum += (strLetter[i] - 'A' + 1);
            }
            return intOutNum;

        }

        public static int getLastCol(Excel.Worksheet Wks)
        {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
        }

        public static int getLastRow(Excel.Worksheet Wks)
        {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        private static void setCursorToWaiting()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlWait;
        }

        private static void setCursorToDefault()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlDefault;
        }


        public static void turnAppSettings(string strDoWhat, Excel.Application xls, bool boolTestCode)
        {
            bool boolOn = true;

            if (strDoWhat == "Off")
                boolOn = false;

            xls.EnableEvents = boolOn;
            xls.ScreenUpdating = boolOn;

            if (boolOn)
                 xls.Calculation = XlCalculation.xlCalculationAutomatic;
            else
                xls.Calculation = XlCalculation.xlCalculationManual;

        }

        public static void formatCells(Excel.Worksheet Wks, decimal intStartRow, decimal intLastRow, decimal intStartColum, decimal intNumColums, string strDoWhat)
        {

            // need to loop for cols - 1gvb2
            decimal intRowCount;

            for (intRowCount = intStartRow; intRowCount <= intLastRow; intRowCount++)
            {
                for (decimal intColCount = intStartColum; intColCount <= intNumColums; intColCount++)
                {
                    Excel.Range xlCell;
                    xlCell = Wks.Cells[intRowCount, intColCount];
                    xlCell.Font.Color = ColorTranslator.FromHtml("Black");
                    xlCell.Interior.Color = ColorTranslator.FromHtml("White");
                    xlCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlCell.Borders.Color = ColorTranslator.ToOle(Color.LightGray); ;
                    xlCell.Borders.Weight = 2d;
                }
            }

        }

        public static void clearFormattingRange(Excel.Worksheet Wks)
        {
            // as I have the worksheet it can all be done here            
            string strRange = "A1:" + getExcelColumnLetter(getLastCol(Wks)) + getLastRow(Wks);

            // this will format the entire range supplied
            Excel.Range xlCell;
            xlCell = Wks.get_Range(strRange);
            xlCell.Font.Color = ColorTranslator.FromHtml("Black");
            xlCell.Interior.Color = ColorTranslator.FromHtml("White");
            xlCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlCell.Borders.Color = ColorTranslator.ToOle(Color.LightGray); ;
            xlCell.Borders.Weight = 2d;

        }


        public static void addValidationToColumn(Excel.Worksheet Wks, string strCol, decimal intStartRow, decimal intEndRow, string strFormula)
        {
            Excel.Range xlCell;
            string strRange = strCol + intStartRow.ToString() + ":" + strCol + intEndRow.ToString();

            xlCell = Wks.get_Range(strRange);
            xlCell.Validation.Delete();
            xlCell.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, Formula1: strFormula);

            xlCell.Validation.InCellDropdown = true;
            xlCell.Validation.ErrorTitle = "Error in Validation";
            xlCell.Validation.ErrorMessage = "Please select value from list";

            // do I want to release this?
            Marshal.ReleaseComObject(xlCell);
            // Marshal.ReleaseComObject(Wks);

        }


        public static string createFormula(string strDeltaCol, decimal intRowStart, decimal intRowEnd)
        {
            string strDeltaRange, strSumString;

            strDeltaRange = strDeltaCol + intRowStart.ToString() + ":" + strDeltaCol + intRowEnd.ToString();
            strSumString = "=SUM(" + strDeltaRange + ")";

            return strSumString;

        }

        public static bool dayCheck(string strValue)
        {
            bool boolRetVal = false;

            if (strValue.Length > 0)
            {
                if (strValue.Length == 19)
                {
                    //if (CommonExcelClasses.IsDate(strValue))
                    if (CommonExcelClasses.IsDate(strValue))
                    {
                        string strDayOfWeek = FormatDate(DateTime.Parse(strValue), "dddd");
                        boolRetVal = (strDayOfWeek == "Monday" || strDayOfWeek == "Tuesday" || strDayOfWeek == "Wednesday" || strDayOfWeek == "Thursday" || strDayOfWeek == "Friday");
                    }
                }
            }

            return boolRetVal;
        }

        public static string FormatDate(DateTime dateTime, string strFormat)
        {
            //return dateTime.ToString("dd/MM/yyyy ");

            if (dateTime == SqlDateTime.MinValue.Value)
                return string.Empty;
            else
                return dateTime.ToString(strFormat);
        }

        /// <summary>
        /// tbd - 1gvb2
        /// </summary>
        /// <param name="Wks"></param>
        /// <param name="intSourceRow"></param>
        /// <param name="strDoWhat"></param>
        /// <param name="intNoCheckCols"></param>
        /// <param name="clrWhichColourFore"></param>
        /// <param name="clrWhichColourBack"></param>
        /// <param name="boolTestCode"></param>
        public static void colourCells(Excel.Worksheet Wks, decimal intSourceRow, string strDoWhat, decimal intNoCheckCols, Color clrWhichColourFore, Color clrWhichColourBack, bool boolTestCode)
        {
            int intStartColumToCheck = 1;

            Excel.Range xlCell;

            for (int intColCount = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
            {
                xlCell = Wks.Cells[intSourceRow, intColCount];

                if (strDoWhat == "Error" || strDoWhat == "Colour")
                {
                    // xlCell = Wks.Cells[intSourceRow, intColCount];
                    xlCell.Font.Color = ColorTranslator.ToOle(clrWhichColourFore);
                    xlCell.Interior.Color = ColorTranslator.ToOle(clrWhichColourBack);
                    xlCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlCell.Borders.Color = ColorTranslator.ToOle(Color.LightGray); ;
                    xlCell.Borders.Weight = 2d;
                }
                else
                {
                    // Wks.Cells[intSourceRow, intColCount].Value2 = null;
                    xlCell.Value2 = null;
                }

                Marshal.ReleaseComObject(xlCell);
            }

            // Marshal.Release(xlCell);
            // Release our resources.
            // Marshal.ReleaseComObject(Wks);
            // Marshal.ReleaseComObject(workbooks);
            // Marshal.FinalReleaseComObject(xlCell);

        }

    }
}
