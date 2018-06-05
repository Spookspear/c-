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


namespace ToolbarOfFunctions_CommonClasses
{
    public class CommonExcelClasses
    {

//         public static string strFilename = "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml";

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

            }

            MessageBox.Show(strMessage, strCaption, MessageBoxButtons.OK, whichIcon);
            // MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question
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

        public static string getExcelColumnName(int columnNumber)
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

        /*
        public static string readProperty(string strWhichProperty)
        {
            // load data
            if (File.Exists(strFilename))
            {
                XmlSerializer xs = new XmlSerializer(typeof(InformationFromSettingsForm));
                FileStream read = new FileStream(strFilename, FileMode.Open, FileAccess.Read, FileShare.Read);
                InformationFromSettingsForm info = (InformationFromSettingsForm)xs.Deserialize(read);
                if (strWhichProperty == "strCompareOrColour")
                {
                    string strRetVal = info.Differences;
                    read.Close();
                    return strRetVal;
                }

            }

            return "Could not Find";
    }
    */

    }
}
