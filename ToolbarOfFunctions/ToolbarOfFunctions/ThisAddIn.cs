#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;


namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        // Clear the workbook
        public void zapWorksheet(Excel.Workbook Wkb)
        {
            Excel.Worksheet Wks;
            Excel.Range xlCell;

            Wks = Wkb.ActiveSheet;

            int intFirstRow = 2;
            int intLastRow = getLastRow(Wks);
            xlCell = Wks.get_Range("A" + intFirstRow + ":A" + intLastRow);

            if (Wks.Name != "InternalParameters")
            {
                if (intLastRow > intFirstRow)
                {
                    xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);
                }
            }
            else
            {
                MsgBox("Cannot run in worksheet: InternalParameters");

            }

        }



        public void MsgBox(string strMessage)
        {
            MessageBox.Show(strMessage);
        }


        public void readFolders(Excel.Workbook Wkb)
        {
            MsgBox("readFolders - code goes here");
        }


        public void compareSheets(Excel.Workbook Wkb)
        {
            MsgBox("compareSheets - code goes here");
        }


        public void deleteBlankLines(Excel.Workbook Wkb)
        {
            Excel.Worksheet Wks;
            Wks = Wkb.ActiveSheet;

            Excel.Range xlCell;

            int intFirstRow = 2;
            int intColScore = 0;

            int intLastRow = getLastRow(Wks);
            int intLastCol = getLastCol(Wks);

            if (Wks.Name != "InternalParameters")
            {
                // loop along looking for data
                for (int intRows = intLastRow; intRows >= intFirstRow; intRows--)
                {
                    Console.WriteLine(intRows);

                    for (int intCols = 1; intCols <= intLastCol; intCols++)
                    {
                        Console.WriteLine(intCols);

                        if (isEmpty(Wks.Cells[intRows, intCols]))
                        {
                            intColScore++;
                        }

                    }

                    if (intColScore == intLastCol)
                    {
                        string strRange = "A" + intRows + ":A" + intRows;
                        xlCell = Wks.get_Range(strRange);
                        xlCell.EntireRow.Delete(Excel.XlDirection.xlUp);


                    }

                    // re initilise the score
                    intColScore = 0;

                }

            }
            else
            {
                MsgBox("Cannot run in worksheet: InternalParameters");
            }

            // MsgBox("Finshed ...");

        }

        private bool isEmpty(Excel.Range xlCell)
        {
            bool boolRetVal = false;

            if (xlCell == null || xlCell.Value2 == null || xlCell.Value2.ToString() == "")
            {
                boolRetVal = true;
            }

            return boolRetVal;

        }

        // move these to common

        // returns the last col
        private int getLastCol(Excel.Worksheet Wks)
        {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

        }

        // returns the last row
        private int getLastRow(Excel.Worksheet Wks)
        {
            return Wks.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

        }

        private static void setCursorToWaiting()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlWait;
        }
        private static void SetCursorToDefault()
        {
            Excel.Application application = Globals.ThisAddIn.Application;
            application.Cursor = Excel.XlMousePointer.xlDefault;
        }

        private bool WorksheetExist(Excel.Workbook Wkb, string strSheetName)
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

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
