#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

using ToolbarOfFunctions_CommonClasses;
using System.DirectoryServices.AccountManagement;

namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {

        public void pingServersIntoWorksheet(Excel.Application xls)
        {

            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolTestCode = myData.TestCode;

            #endregion

            #region [Declare and instantiate variables for worksheet/book]
            // get worksheet name
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet Wks;   // get current sheet

            Wks = Wkb.ActiveSheet;

            string strMessage = "Ping Servers in Column: " + CommonExcelClasses.getExcelColumnLetter((int)myData.ColPingRead);
            strMessage = strMessage + "and write results into column: " + CommonExcelClasses.getExcelColumnLetter((int)myData.ColPingWrite);


            #endregion

            #region [Ask to display a Message?]
            DialogResult dlgResult = DialogResult.Yes;

            if (boolDisplayInitialMessage)
            {
                if (booltimeTaken)
                    strMessage = strMessage + LF + " and display the time taken";

                dlgResult = MessageBox.Show(strMessage + "?", "Active Directory", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            }

            #endregion


            #region [Start of work]
            if (dlgResult == DialogResult.Yes)
            {

                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("Off", xls, myData.TestCode);

                DateTime dteStart = DateTime.Now;

                // do work


                if (boolTurnOffScreen)
                    CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);


                #region [Display Complete Message]
                if (boolDisplayCompleteMessage)
                {
                    strMessage = "";
                    strMessage = strMessage + "Complete ...";

                    if (booltimeTaken)
                    {
                        DateTime dteEnd = DateTime.Now;
                        int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                        strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                    }

                    CommonExcelClasses.MsgBox(strMessage);          // localisation?
                }
                #endregion
            }

            #endregion


        }

    }

}
