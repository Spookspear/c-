#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;

using System.Xml.Linq;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using System.ComponentModel;
using System.Data;

// using ToolbarOfFunctions;
using ToolbarOfFunctions_CommonClasses;


namespace ToolbarOfFunctions_Graveyard
{
    class classGraveYard
    {

        // holds useful code
        // Excel.Application excel = new Excel.Application();

        private static void graveYard()
        {
            // Excel.Worksheet activeWorksheet = null;
            // int cnt = 0;
            //foreach (Excel.Range element in range.Cells)            {
            //    if (element.Value2 != null)
            //    {
            //        cnt = cnt + 1;
            //    }
            //    System.Console.WriteLine(cnt);
            //}
            //MessageBox.Show(cnt.ToString());

            // Excel.Worksheet activeWorksheet;

            // MessageBox.Show(activeWorksheet.Name.ToLower());
            // string activeWorksheet = Wkb.Sheets[0].name;
            //activeWorksheet = Wkb.Sheets[1];
            //MessageBox.Show(activeWorksheet._CodeName);
            //WorksheetExist(Wkb, "Sheet1");

            // Excel.Range cell = activeWorksheet.get_Range(x.column + x.row);
            // string activeWorksheetName = activeWorksheet.Name;
            // MessageBox.Show(activeWorksheet.Name);
            // MessageBox.Show(activeWorksheetName);
            // setCursorToWaiting();

        }

        private static void deleteEmptyRowsCols(Excel.Worksheet worksheet)
        {
            Excel.Range targetCells = worksheet.UsedRange;
            object[,] allValues = (object[,])targetCells.Cells.Value;
            int totalRows = targetCells.Rows.Count;
            int totalCols = targetCells.Columns.Count;

            List<int> emptyRows = getEmptyRows_graveYard(allValues, totalRows, totalCols);
            List<int> emptyCols = getEmptyCols(allValues, totalRows, totalCols);

            // now we have a list of the empty rows and columns we need to delete
            deleteRows_graveYard(emptyRows, worksheet);
            deleteCols(emptyCols, worksheet);
        }



        private static void deleteRows_graveYard(List<int> rowsToDelete, Excel.Worksheet worksheet)
        {
            // the rows are sorted high to low - so index's wont shift
            foreach (int rowIndex in rowsToDelete)
            {
                worksheet.Rows[rowIndex].Delete();
            }
        }


        private static List<int> getEmptyRows_graveYard(object[,] allValues, int totalRows, int totalCols)
        {
            List<int> emptyRows = new List<int>();

            for (int i = 1; i < totalRows; i++)
            {
                if (IsRowEmpty_graveYard(allValues, i, totalCols))
                {
                    emptyRows.Add(i);
                }
            }
            // sort the list from high to low
            return emptyRows.OrderByDescending(x => x).ToList();
        }


        private static bool IsRowEmpty_graveYard(object[,] allValues, int rowIndex, int totalCols)
        {
            for (int i = 1; i < totalCols; i++)
            {
                if (allValues[rowIndex, i] != null)
                {
                    return false;
                }
            }
            return true;
        }

        private static List<int> getEmptyCols(object[,] allValues, int totalRows, int totalCols)
        {
            List<int> emptyCols = new List<int>();

            for (int i = 1; i < totalCols; i++)
            {
                if (IsColumnEmpty(allValues, i, totalRows))
                {
                    emptyCols.Add(i);
                }
            }
            // sort the list from high to low
            return emptyCols.OrderByDescending(x => x).ToList();
        }


        private static void deleteCols(List<int> colsToDelete, Excel.Worksheet worksheet)
        {
            // the cols are sorted high to low - so index's wont shift
            foreach (int colIndex in colsToDelete)
            {
                worksheet.Columns[colIndex].Delete();
            }
        }

        private static void deleteTopEmptyRows(Excel.Worksheet worksheet, int startRow)
        {
            for (int i = 0; i < startRow - 1; i++)
            {
                worksheet.Rows[1].Delete();
            }
        }


        private static bool IsColumnEmpty(object[,] allValues, int colIndex, int totalRows)
        {
            for (int i = 1; i < totalRows; i++)
            {
                if (allValues[i, colIndex] != null)
                {
                    return false;
                }
            }
            return true;
        }

        private static void removeEmptyTopRowsAndLeftCols(Excel.Worksheet worksheet, Excel.Range usedRange)
        {
            string addressString = usedRange.Address.ToString();
            int rowsToDelete = getNumberOfTopRowsToDelete(addressString);
            deleteTopEmptyRows(worksheet, rowsToDelete);
            int colsToDelete = getNumberOfLeftColsToDelte(addressString);
            deleteLeftEmptyColumns(worksheet, colsToDelete);
        }

        private static void deleteLeftEmptyColumns(Excel.Worksheet worksheet, int colCount)
        {
            for (int i = 0; i < colCount - 1; i++)
            {
                worksheet.Columns[1].Delete();
            }
        }

        private static int getNumberOfTopRowsToDelete(string address)
        {
            string[] splitArray = address.Split(':');
            string firstIndex = splitArray[0];
            splitArray = firstIndex.Split('$');
            string value = splitArray[2];
            int returnValue = -1;
            if ((int.TryParse(value, out returnValue)) && (returnValue >= 0))
                return returnValue;
            return returnValue;
        }

        private static int getNumberOfLeftColsToDelte(string address)
        {
            string[] splitArray = address.Split(':');
            string firstindex = splitArray[0];
            splitArray = firstindex.Split('$');
            string value = splitArray[1];
            return parseColHeaderToIndex(value);
        }

        private static int parseColHeaderToIndex(string colAdress)
        {
            int[] digits = new int[colAdress.Length];
            for (int i = 0; i < colAdress.Length; ++i)
            {
                digits[i] = Convert.ToInt32(colAdress[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        public void compareSheets_graveYard(Excel.Workbook Wkb)
        {
            Excel.Worksheet Wks1;               // get current sheet
            Excel.Worksheet Wks2;               // get sheet next door
            string strClearOrColour = "Colour";

            Wks1 = Wkb.ActiveSheet;
            Wks2 = Wkb.Sheets[Wks1.Index + 1];

            DialogResult dlgResult = MessageBox.Show("Compare: Worksheet: " + Wks1.Name + " against: " + Wks2.Name + " and " + strClearOrColour + " ones which are the same?", "Compare Sheets", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (dlgResult == DialogResult.Yes)
            {

                int intTargetRow = 0;
                int intStartRow = 2;

                // how may columns to check ?
                int intNoCheckCols = 5;             // for later loop?
                int intStartColumToCheck = 1;
                int intColScore = 1;

                string strValue1 = "";

                int intSheetLastRow1 = CommonExcelClasses.getLastRow(Wks1);
                int intSheetLastRow2 = CommonExcelClasses.getLastRow(Wks2);

                for (int intSourceRow = intStartRow; intSourceRow <= intSheetLastRow1; intSourceRow++)
                {
                    // read in vlaue from sheet 
                    // maybe I should ready all into arrayS?

                    strValue1 = Wks1.Cells[intSourceRow, intStartColumToCheck].Value;

                    intTargetRow = CommonExcelClasses.searchForValue(Wks2, strValue1, intStartColumToCheck);

                    if (intTargetRow > 0)
                    {
                        //  start from correct column
                        for (int intColCount = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
                        {
                            // Compare cells directly
                            if (Wks1.Cells[intSourceRow, intColCount].Value == Wks2.Cells[intTargetRow, intColCount].Value)
                            {
                                intColScore++;
                            }

                        }

                        // Score system = if all the same then can blue it
                        if (intColScore == intNoCheckCols)
                        {
                            for (int intColCount = intStartColumToCheck; intColCount <= intNoCheckCols; intColCount++)
                            {
                                if (strClearOrColour == "Colour")
                                {
                                    Wks1.Cells[intSourceRow, intColCount].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue); ;
                                }
                            }
                        }

                        intColScore = 1;
                    }
                }
            }
        }

    }

    // check for exit
    // Excel.EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt;
    // Excel.XlEnableCancelKey key = XlEnableCancelKey.xlErrorHandler;
    // Globals.ThisAddIn.Application.SendKeys("{ESC}");

}
