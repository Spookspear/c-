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

using System.Net.NetworkInformation;

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

            int intStartRow = (int)myData.ComparingStartRow;
            int intServerColumn = (int)myData.ColPingRead;
            int intWriteColumn = (int)myData.ColPingWrite;

            #endregion
            int intSourceRow = intStartRow;

            try
            {
                #region [Declare and instantiate variables for worksheet/book]

                Excel.Workbook Wkb = xls.ActiveWorkbook;
                Excel.Worksheet Wks;
                Wks = Wkb.ActiveSheet;      // get current sheet

                string strMessage = "Ping Servers in Column: " + CommonExcelClasses.getExcelColumnLetter((int)myData.ColPingRead);
                strMessage = strMessage + " and write results into column: " + CommonExcelClasses.getExcelColumnLetter((int)myData.ColPingWrite);

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
                    int intSheetLastRow = CommonExcelClasses.getLastRow(Wks);
                    string strServer = "";

                    // do work
                    for (intSourceRow = intStartRow; intSourceRow <= intSheetLastRow; intSourceRow++)
                    {
                        // read in vlaue from sheet 
                        // maybe I should ready all into arrays - maybe later?

                        if (!CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow, intServerColumn]))
                        {
                            strServer = Wks.Cells[intSourceRow, intServerColumn].Value;

                            if (!boolTestCode)
                            {
                                if (CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow, intServerColumn+1]))
                                {
                                    Wks.Cells[intSourceRow, intWriteColumn].Value = PingHost(strServer);
                                    Wks.Cells[intSourceRow, intWriteColumn + 1].Value = "PingHost(strServer)";
                                }
                            }
                            else
                            {
                                if (CommonExcelClasses.isEmptyCell(Wks.Cells[intSourceRow, intServerColumn + 1]))
                                {
                                    Wks.Cells[intSourceRow, intWriteColumn].Value = PingHost2(strServer);
                                    Wks.Cells[intSourceRow, intWriteColumn + 1].Value = "PingHost2(strServer)";
                                }
                            }

                        }

                    }


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
            catch (System.Exception excpt)
            {
                CommonExcelClasses.MsgBox("Issues around row" + intSourceRow.ToString(), "Error");
                Console.WriteLine(excpt.Message);
            }

        }


        public static string PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = null;
            string strRetVal = "";
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;

                if (pingable)
                {

                    strRetVal  = reply.RoundtripTime.ToString() + " ms";
                }

            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }


            return strRetVal;
        }

        public static string PingHost2(string strServer)
        {
            Ping p = new Ping();
            PingReply r;
            string strRetVal = "";

            r = p.Send(strServer);

            if (r.Status == IPStatus.Success) {
                strRetVal =  "Ping to " + strServer.ToString() + "[" + r.Address.ToString() + "]" + " Successful Response delay = " + r.RoundtripTime.ToString() + " ms";
            }

            return strRetVal.ToString();
        }

    }

}
