﻿#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
// using ToolbarOfFunctions = ToolbarOfFunctions.ThisAddIn;
using ToolbarOfFunctions_CommonClasses;
using ToolbarOfFunctions;



using System.Xml.Serialization;
using System.IO;

namespace ToolbarOfFunctions
{
    public partial class ExcelRibbon
    {
        public bool boolDisplayMessage, boolLargeButton, boolHideText;
        public string strCompareOrColour;

        frmSettings frmSettings = new frmSettings();
        InformationForSettingsForm myData = new InformationForSettingsForm();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            myData = SaveXML.LoadData();
            // strCompareOrColour = SaveXML.readProperty("strCompareOrColour");
            CommonExcelClasses.ButtonUpdateLabel(btnCompareSheets, "Compare (" + myData.CompareOrColour + ")");

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
            Globals.ThisAddIn.compareSheets(Globals.ThisAddIn.Application);
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
            DialogResult dr = frmSettings.ShowDialog();

            if (dr == DialogResult.OK)
            {
                boolDisplayMessage = frmSettings.chkProduceInitialMessageBox.Checked;
                boolLargeButton = frmSettings.chkLargeButtons.Checked;
                boolHideText = frmSettings.chkHideText.Checked;

                CommonExcelClasses.ButtonSetSize(btnSettings, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnReadFolders, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnCompareSheets, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnZap, boolLargeButton);
                CommonExcelClasses.SplitButtonSetSize(splitButtonDeleteLines, boolLargeButton);

                CommonExcelClasses.ButtonSetSize(btnDeleteBlankLinesA, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnDeleteBlankLinesB, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnDeleteBlankLinesC, boolLargeButton);

                CommonExcelClasses.ButtonSetSize(btnDealWithSingleDuplicates, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnDealWithManyDuplicates, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnLoadADGroupIntoSpreadsheet, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnLoadADGroupIntoSpreadsheetActiveCell, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnReadUsersGroupMembership, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnReadUsers, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnWriteTimeSheet, boolLargeButton);
                CommonExcelClasses.ButtonSetSize(btnPingServers, boolLargeButton);


                if (boolHideText) {
                    CommonExcelClasses.ButtonUpdateLabel(btnSettings, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadFolders, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnCompareSheets, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnZap, "");
                    CommonExcelClasses.SplitButtonUpdateLabel(splitButtonDeleteLines, "");

                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesA, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesB, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesC, "");

                    CommonExcelClasses.ButtonUpdateLabel(btnDealWithSingleDuplicates, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnDealWithManyDuplicates, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnLoadADGroupIntoSpreadsheet, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnLoadADGroupIntoSpreadsheetActiveCell, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadUsersGroupMembership, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadUsers, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnWriteTimeSheet, "");
                    CommonExcelClasses.ButtonUpdateLabel(btnPingServers, "");

                } else {

                    CommonExcelClasses.ButtonUpdateLabel(btnSettings, "Settings");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadFolders, "Read Folders");
                    
                    // reload - 1gvb2
                    myData = SaveXML.LoadData();

                    CommonExcelClasses.ButtonUpdateLabel(btnCompareSheets, "Compare (" + myData.CompareOrColour + ")");
                    CommonExcelClasses.ButtonUpdateLabel(btnZap, "Zap Worksheet");
                    CommonExcelClasses.SplitButtonUpdateLabel(splitButtonDeleteLines, "Delete Blank Lines");
                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesA, "Mode: A");
                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesB, "Mode: B");
                    CommonExcelClasses.ButtonUpdateLabel(btnDeleteBlankLinesC, "Mode: C");

                    CommonExcelClasses.ButtonUpdateLabel(btnDealWithSingleDuplicates, "Duplicates (Cols: Single):");
                    CommonExcelClasses.ButtonUpdateLabel(btnDealWithManyDuplicates, "Duplicates (Cols: Many)");
                    CommonExcelClasses.ButtonUpdateLabel(btnLoadADGroupIntoSpreadsheet, "AD Group Members");
                    CommonExcelClasses.ButtonUpdateLabel(btnLoadADGroupIntoSpreadsheetActiveCell, "AD Members - Active Cell");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadUsersGroupMembership, "Users AD Membership");
                    CommonExcelClasses.ButtonUpdateLabel(btnReadUsers, "Details from AD Name");
                    CommonExcelClasses.ButtonUpdateLabel(btnWriteTimeSheet, "Update timesheet");
                    CommonExcelClasses.ButtonUpdateLabel(btnPingServers, "Ping Servers");

                }
            }

        }

        public void btnDealWithSingleDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.dealWithSingleDuplicates(Globals.ThisAddIn.Application.ActiveWorkbook);

        }

    }

}
