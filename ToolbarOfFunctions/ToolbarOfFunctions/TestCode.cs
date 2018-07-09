using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
// using Microsoft.Office.Interop.Excel;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using System.ComponentModel;
using System.Data;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Tools.Ribbon;

using ToolbarOfFunctions_CommonClasses;
using ToolbarOfFunctions_MyConstants;
using System.Runtime.InteropServices;

using System.Data.SqlTypes;


namespace ToolbarOfFunctions
{
    class TestCode
    {


        private void button1_Click(object sender, EventArgs e)

        {

            DataTable dt1 = new DataTable();

            dt1.Columns.Add("Name");



            dt1.Rows.Add("A");

            dt1.Rows.Add("B");

            dt1.Rows.Add("C");

            dt1.Rows.Add("D");

            dt1.Rows.Add("E");



            Excel.Application oXL = new Excel.Application();

            oXL.SheetsInNewWorkbook = 2;



            Excel.Workbook oWB = oXL.Workbooks.Add();



            Excel.Worksheet oMasterSheet = oWB.Worksheets["sheet2"];

            oMasterSheet.Name = "MasterData";



            Excel.Worksheet oSheet = oWB.Worksheets["sheet1"];

            oSheet.Name = "UserData";



            SetExcelMasterData(dt1, oMasterSheet, "Name", "Name", 1, oSheet, "A", "A");



            string sFileName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "ExcelFile.xlsx";



            oWB.SaveAs(sFileName);

            oWB.Close(Type.Missing);



            Application.DoEvents();

            Process showXL = Process.Start(sFileName);

            Application.DoEvents();

        }



        private void SetExcelMasterData(DataTable dt, Excel.Worksheet oMasterSheet, string sMasterHeader, string sField, int iCol, Excel.Worksheet oSheet, string sCell, string sMasterCell)

        {

            oMasterSheet.Cells.Item[1, iCol] = sMasterHeader;

            for (int i = 0; i < dt.Rows.Count; i++)

            {

                oMasterSheet.Cells.Item[i + 2, iCol] = dt.Rows[i][sField];

            }



            oSheet.Range[sCell + "2"].EntireColumn.Validation.Add(Excel.XlDVType.xlValidateList, (Excel.XlDVAlertStyle.xlValidAlertStop), Operator: Excel.XlFormatConditionOperator.xlBetween, Formula1: "=" + oMasterSheet.Name + "!$" + sMasterCell + "$2:$" + sMasterCell + "$" + (dt.Rows.Count + 1).ToString());

            oSheet.Range[sCell + "2"].EntireColumn.Validation.InCellDropdown = true;

            oSheet.Range[sCell + "2"].EntireColumn.Validation.ErrorTitle = "Error in Validation";

            oSheet.Range[sCell + "2"].EntireColumn.Validation.ErrorMessage = "Please select value from list";

        }


    }
}
