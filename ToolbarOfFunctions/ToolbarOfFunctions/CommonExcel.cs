#pragma warning disable IDE1006 // Naming Styles

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
    public static class CommonExcelClasses
    {

        const int _ROW1 = 1;
        const int _ROW2 = 3;

        const int _COL1 = 0;
        const int _COL2 = 2;


        // exstensability
        public static void SwitchToBoldRegularChkBox(this System.Windows.Forms.CheckBox c)
        {
            if (c.Font.Style != FontStyle.Bold)
                c.Font = new System.Drawing.Font(c.Font, FontStyle.Bold);
            else
                c.Font = new System.Drawing.Font(c.Font, FontStyle.Regular);
        }


        public static void SwtichToBoldRegularTextBox(this System.Windows.Forms.TextBox c)
        {
            if (c.Font.Style != FontStyle.Bold)
                c.Font = new System.Drawing.Font(c.Font, FontStyle.Bold);
            else
                c.Font = new System.Drawing.Font(c.Font, FontStyle.Regular);
        }

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
            } else {
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


        public static bool WorksheetExist(Excel.Workbook Wkb, string strSheetName)
        {
            bool found = false;

            foreach (Excel.Worksheet Wks in Wkb.Sheets)
            {

                // MsgBox("inside WorksheetExist() - ws Name " + Wks.Name.ToLower());

                if (Wks.Name.ToLower() == strSheetName.ToLower())
                {
                    found = true;
                    break;
                }
            }

            return found;
        }

        /// <summary>
        /// Clear the workbook
        /// </summary>
        public static void zapWorksheet(Excel.Worksheet Wks, int intFirstRow = 2)
        {
            Excel.Range xlCell;

            // Wks = Wkb.ActiveSheet;

            int intLastRow = CommonExcelClasses.getLastRow(Wks);
            xlCell = Wks.get_Range("A" + intFirstRow + ":A" + intLastRow);

            if (Wks.Name != "InternalParameters")
            {
                if (intLastRow > intFirstRow)
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);

            } else {

                CommonExcelClasses.MsgBox("Cannot run in worksheet: InternalParameters", "Error");
            }

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
            } else {
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


        public static void setCursorToWaiting()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlWait;
        }


        public static void setCursorToDefault()
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

                    // 1gvb2 - this is not tested
                    #region [release memory]
                    Marshal.ReleaseComObject(xlCell);
                    #endregion
                }
            }

        }


        public static void clearFormattingRange(Excel.Worksheet Wks)
        {

            #region [Define range dynamically]
            string strRange = "A1:" + getExcelColumnLetter(getLastCol(Wks)) + getLastRow(Wks);
            #endregion

            #region [format the entire range supplied]
            Excel.Range xlCell;
            xlCell = Wks.get_Range(strRange);
            xlCell.Font.Color = ColorTranslator.FromHtml("Black");
            xlCell.Interior.Color = ColorTranslator.FromHtml("White");
            xlCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlCell.Borders.Color = ColorTranslator.ToOle(Color.LightGray); ;
            xlCell.Borders.Weight = 2d;
            #endregion

            #region [release memory]
            Marshal.ReleaseComObject(xlCell);
            #endregion
        }

                public static void sortSheet(Excel.Worksheet Wks, int intColNo)
        {
            // this will sort the sheet on the 1st column
            #region [Define range dynamically]
            string strRange = "A2:" + getExcelColumnLetter(getLastCol(Wks)) + getLastRow(Wks);
            #endregion


            Excel.Range Sheet = Wks.get_Range(strRange);
            Sheet.Sort(
                Sheet.Columns[intColNo], Excel.XlSortOrder.xlAscending,
                Type.Missing, Type.Missing, Excel.XlSortOrder.xlAscending,
                Type.Missing, Excel.XlSortOrder.xlAscending,
                Excel.XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
                Excel.XlSortOrientation.xlSortColumns,
                Excel.XlSortMethod.xlPinYin,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal
                
                );

        }

        public static void sortSheetWorking(Excel.Worksheet Wks)
        {
            // this will sort the sheet on the 1st column
            #region [Define range dynamically]
            string strRange = "A2:" + getExcelColumnLetter(getLastCol(Wks)) + getLastRow(Wks);
            #endregion


            Excel.Range Sheet = Wks.get_Range(strRange);
            Sheet.Sort(
                Sheet.Columns[1], Excel.XlSortOrder.xlAscending,
                Sheet.Columns[2], Type.Missing, Excel.XlSortOrder.xlAscending,
                Type.Missing, Excel.XlSortOrder.xlAscending,
                Excel.XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
                Excel.XlSortOrientation.xlSortColumns,
                Excel.XlSortMethod.xlPinYin,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal

                );

        }


        public static void addValidationToColumn(Excel.Worksheet Wks, string strCol, decimal intStartRow, decimal intEndRow, string strFormula)
        {

            #region [Define range from passed in parameters]
            string strRange = strCol + intStartRow.ToString() + ":" + strCol + intEndRow.ToString();
            #endregion

            #region [format the entire range supplied]
            Excel.Range xlCell;
            xlCell = Wks.get_Range(strRange);
            xlCell.Validation.Delete();
            xlCell.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, Formula1: strFormula);

            xlCell.Validation.InCellDropdown = true;
            xlCell.Validation.ErrorTitle = "Error in Validation";
            xlCell.Validation.ErrorMessage = "Please select value from list";
            #endregion

            #region [release memory]
            Marshal.ReleaseComObject(xlCell);
            #endregion

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


        public static void colourCells(Excel.Worksheet Wks, decimal intSourceRow, string strDoWhat, decimal intNoCheckCols, Color clrWhichColourFore, Color clrWhichColourBack, bool boolTestCode)
        {

            /*
            #region [Declare and instantiate variables for process]
            InformationForSettingsForm myData = new InformationForSettingsForm();

            // dont need to pass in can read diretly from the XML
            // 1gvb2

            myData = myData.LoadMyData();               // read data from settings file

            string strCompareOrColour = myData.CompareOrColour;
            Color clrColourFore_Found = ColorTranslator.FromHtml(myData.ColourFore_Found);
            Color clrColourFore_NotFound = ColorTranslator.FromHtml(myData.ColourFore_NotFound);
            bool boolFontBold_Found = myData.ColourBold_Found;
            bool boolFontBold_NotFound = myData.ColourBold_NotFound;
            Color clrColourBack_Found = ColorTranslator.FromHtml(myData.ColourBack_Found);
            Color clrColourBack_NotFound = ColorTranslator.FromHtml(myData.ColourBack_NotFound);
            int intStartRow = (int)myData.ComparingStartRow;
            bool boolTestCode = myData.TestCode;
            #endregion
            */

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
                } else {
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



        public static DateTime getFileDate(string strFileName)
        {
            FileInfo oFileInfo = new FileInfo(strFileName);
            FileVersionInfo oFileVersionInfo = FileVersionInfo.GetVersionInfo(strFileName);

            return oFileInfo.LastWriteTime;

        }


        public static bool checkEmptyRange(Worksheet Wks, string strRange)
        {
            int[] intArrCoords = getCoordsFromRange(strRange);
            int intNotEmpty = 0;

            // loop along the Rows
            for (int x = intArrCoords[_ROW1]; x <= intArrCoords[_ROW2]; x++)
            {
                // loop along the Cols
                for (int y = intArrCoords[_COL1]; y <= intArrCoords[_COL2]; y++)
                {
                    if (!CommonExcelClasses.isEmptyCell(Wks.Cells[x, y]))
                    {
                        intNotEmpty++;
                        break;
                    }
                }

                if (intNotEmpty > 0)
                    break;

            }

            return (intNotEmpty == 0);


        }

        private static int[] getCoordsFromRange(string strRange)
        {

            int intSPos = strRange.IndexOf(":");

            string strColAddr1 = strRange.Substring(0, intSPos);
            string strColAddr2 = strRange.Substring(intSPos + 1);

            // A20:F20 but could be AA20:FF20 - so now need to get 1st occurance of a number
            int intIdx = strColAddr1.IndexOfAny("0123456789".ToCharArray());
            int intColAddrLtr1 = CommonExcelClasses.getExcelColumnNumber(strColAddr1.Substring(0, intIdx));
            int intRowAddrNum1 = Convert.ToInt32(strColAddr1.Substring(intIdx));

            intIdx = strColAddr2.IndexOfAny("0123456789".ToCharArray());
            int intColAddrLtr2 = CommonExcelClasses.getExcelColumnNumber(strColAddr2.Substring(0, intIdx));
            int intRowAddrNum2 = Convert.ToInt32(strColAddr2.Substring(intIdx));

            int[] intArrReturn = { intColAddrLtr1, intRowAddrNum1, intColAddrLtr2, intRowAddrNum2 };

            return intArrReturn;

        }


        // exstensability
        public static void SwitchMainOrAdditional(this RiggingLinesDS c)
        {
            if (c.LineOrAdditional != "M")
                c.LineOrAdditional = "M";
            else
                c.LineOrAdditional = "A";
        }

    }



}
