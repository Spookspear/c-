#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ToolbarOfFunctions = ToolbarOfFunctions.ThisAddIn;
using ToolbarOfFunctions_CommonClasses;

using ToolbarOfFunctions;



namespace ToolbarOfFunctions
{
    public partial class ExcelRibbon
    {

        public bool boolDisplayMessage, boolLargeButton;

        // frmSettings frmSettings = new frmSettings();
        // frmSettings frmSettings = default(frmSettings);
        frmSettings frmSettings = new frmSettings();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void btnZap_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.zapWorksheet(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void btnReadFolders_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.readFolders(Globals.ThisAddIn.Application.ActiveWorkbook);
        }       

        private void btnCompareSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.compareSheets(Globals.ThisAddIn.Application.ActiveWorkbook);
        }
        
        private void btnDeleteBlankLinesA_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "A");
        }


        private void btnDeleteBlankLinesB_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "B");

        }

        private void btnDeleteBlankLinesC_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "C");

        }

        private void splitButtonDeleteLines_Click(object sender, RibbonControlEventArgs e)
        {
            btnDeleteBlankLinesB_Click(sender, e);
        }

        public void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            // frmSettings frmSettings = new frmSettings();
            // Globals.ThisAddIn.openSettingsForm(Globals.ThisAddIn.Application.ActiveWorkbook);
            // btnDealWithSingleDuplicates.Label = "Hi";

            frmSettings.ShowDialog();

            boolDisplayMessage = frmSettings.chkProduceMessageBox.Checked;
            boolLargeButton = frmSettings.chkLargeButtons.Checked;

            CommonExcelClasses.ButtonSetSize(btnSettings, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnReadFolders, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnCompareSheets, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnZap, boolLargeButton);
            CommonExcelClasses.SplitButtonSetSize(splitButtonDeleteLines, boolLargeButton);


            CommonExcelClasses.ButtonSetSize(btnDealWithSingleDuplicates, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnDealWithManyDuplicates, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnLoadADGroupIntoSpreadsheet, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnLoadADGroupIntoSpreadsheetActiveCell, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnReadUsersGroupMembership, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnReadUsers, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnWriteTimeSheet, boolLargeButton);
            CommonExcelClasses.ButtonSetSize(btnPingServers, boolLargeButton);


        }

        public void btnDealWithSingleDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.dealWithSingleDuplicates(Globals.ThisAddIn.Application.ActiveWorkbook);

        }
    }

}
