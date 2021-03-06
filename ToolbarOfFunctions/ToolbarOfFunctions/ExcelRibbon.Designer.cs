﻿namespace ToolbarOfFunctions
{
    partial class ExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelRibbon));
            this.MYTOOLBAR = this.Factory.CreateRibbonTab();
            this.customToolbar = this.Factory.CreateRibbonGroup();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnReadFolders = this.Factory.CreateRibbonButton();
            this.btnCompareSheets = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnZap = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.splitButtonDeleteLines = this.Factory.CreateRibbonSplitButton();
            this.btnDeleteBlankLinesA = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLinesB = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLinesC = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnDealWithSingleDuplicates = this.Factory.CreateRibbonButton();
            this.btnDealWithManyDuplicates = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.btnReadUsersGroupMembership = this.Factory.CreateRibbonButton();
            this.btnReadUsersGroupMembershipActiveCell = this.Factory.CreateRibbonButton();
            this.btnLoadADGroupIntoSpreadsheet = this.Factory.CreateRibbonButton();
            this.btnLoadADGroupIntoSpreadsheetActiveCell = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.btnWriteTimeSheet = this.Factory.CreateRibbonButton();
            this.btnPingServers = this.Factory.CreateRibbonButton();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.MYTOOLBAR.SuspendLayout();
            this.customToolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // MYTOOLBAR
            // 
            this.MYTOOLBAR.Groups.Add(this.customToolbar);
            this.MYTOOLBAR.Label = "MYTOOLBAR";
            this.MYTOOLBAR.Name = "MYTOOLBAR";
            // 
            // customToolbar
            // 
            this.customToolbar.Items.Add(this.btnSettings);
            this.customToolbar.Items.Add(this.separator1);
            this.customToolbar.Items.Add(this.btnReadFolders);
            this.customToolbar.Items.Add(this.btnCompareSheets);
            this.customToolbar.Items.Add(this.separator2);
            this.customToolbar.Items.Add(this.btnZap);
            this.customToolbar.Items.Add(this.separator3);
            this.customToolbar.Items.Add(this.splitButtonDeleteLines);
            this.customToolbar.Items.Add(this.separator4);
            this.customToolbar.Items.Add(this.btnDealWithSingleDuplicates);
            this.customToolbar.Items.Add(this.btnDealWithManyDuplicates);
            this.customToolbar.Items.Add(this.separator5);
            this.customToolbar.Items.Add(this.btnReadUsersGroupMembership);
            this.customToolbar.Items.Add(this.btnReadUsersGroupMembershipActiveCell);
            this.customToolbar.Items.Add(this.btnLoadADGroupIntoSpreadsheet);
            this.customToolbar.Items.Add(this.btnLoadADGroupIntoSpreadsheetActiveCell);
            this.customToolbar.Items.Add(this.separator6);
            this.customToolbar.Items.Add(this.btnWriteTimeSheet);
            this.customToolbar.Items.Add(this.btnPingServers);
            this.customToolbar.Label = "Custom Toolbar";
            this.customToolbar.Name = "customToolbar";
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Description = "Settings";
            this.btnSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSettings.Image")));
            this.btnSettings.Label = "Settings";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnReadFolders
            // 
            this.btnReadFolders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadFolders.Description = "Read Folders into Worksheet";
            this.btnReadFolders.Image = ((System.Drawing.Image)(resources.GetObject("btnReadFolders.Image")));
            this.btnReadFolders.Label = "Read Folders";
            this.btnReadFolders.Name = "btnReadFolders";
            this.btnReadFolders.ShowImage = true;
            this.btnReadFolders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadFolders_Click);
            // 
            // btnCompareSheets
            // 
            this.btnCompareSheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCompareSheets.Description = "Compare against the sheet next door";
            this.btnCompareSheets.Image = ((System.Drawing.Image)(resources.GetObject("btnCompareSheets.Image")));
            this.btnCompareSheets.Label = "Compare Sheets";
            this.btnCompareSheets.Name = "btnCompareSheets";
            this.btnCompareSheets.ShowImage = true;
            this.btnCompareSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompareSheets_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnZap
            // 
            this.btnZap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnZap.Description = "Zap Worksheet";
            this.btnZap.Image = ((System.Drawing.Image)(resources.GetObject("btnZap.Image")));
            this.btnZap.Label = "Zap Worksheet";
            this.btnZap.Name = "btnZap";
            this.btnZap.ShowImage = true;
            this.btnZap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnZap_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // splitButtonDeleteLines
            // 
            this.splitButtonDeleteLines.ButtonType = Microsoft.Office.Tools.Ribbon.RibbonButtonType.ToggleButton;
            this.splitButtonDeleteLines.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonDeleteLines.Description = "Mode type";
            this.splitButtonDeleteLines.Image = ((System.Drawing.Image)(resources.GetObject("splitButtonDeleteLines.Image")));
            this.splitButtonDeleteLines.Items.Add(this.btnDeleteBlankLinesA);
            this.splitButtonDeleteLines.Items.Add(this.btnDeleteBlankLinesB);
            this.splitButtonDeleteLines.Items.Add(this.btnDeleteBlankLinesC);
            this.splitButtonDeleteLines.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButtonDeleteLines.Label = "Delete Blank Lines";
            this.splitButtonDeleteLines.Name = "splitButtonDeleteLines";
            this.splitButtonDeleteLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButtonDeleteLines_Click);
            // 
            // btnDeleteBlankLinesA
            // 
            this.btnDeleteBlankLinesA.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteBlankLinesA.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesA.Image")));
            this.btnDeleteBlankLinesA.Label = "Mode: A";
            this.btnDeleteBlankLinesA.Name = "btnDeleteBlankLinesA";
            this.btnDeleteBlankLinesA.ShowImage = true;
            this.btnDeleteBlankLinesA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesA_Click);
            // 
            // btnDeleteBlankLinesB
            // 
            this.btnDeleteBlankLinesB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteBlankLinesB.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesB.Image")));
            this.btnDeleteBlankLinesB.Label = "Mode: B";
            this.btnDeleteBlankLinesB.Name = "btnDeleteBlankLinesB";
            this.btnDeleteBlankLinesB.ScreenTip = "Fastest";
            this.btnDeleteBlankLinesB.ShowImage = true;
            this.btnDeleteBlankLinesB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesB_Click);
            // 
            // btnDeleteBlankLinesC
            // 
            this.btnDeleteBlankLinesC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteBlankLinesC.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesC.Image")));
            this.btnDeleteBlankLinesC.Label = "Mode: C";
            this.btnDeleteBlankLinesC.Name = "btnDeleteBlankLinesC";
            this.btnDeleteBlankLinesC.ScreenTip = "Slowest";
            this.btnDeleteBlankLinesC.ShowImage = true;
            this.btnDeleteBlankLinesC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesC_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // btnDealWithSingleDuplicates
            // 
            this.btnDealWithSingleDuplicates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDealWithSingleDuplicates.Description = "Duplicates (Cols: Single):";
            this.btnDealWithSingleDuplicates.Image = ((System.Drawing.Image)(resources.GetObject("btnDealWithSingleDuplicates.Image")));
            this.btnDealWithSingleDuplicates.Label = "Duplicates (Cols: Single):";
            this.btnDealWithSingleDuplicates.Name = "btnDealWithSingleDuplicates";
            this.btnDealWithSingleDuplicates.ShowImage = true;
            this.btnDealWithSingleDuplicates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDealWithSingleDuplicates_Click);
            // 
            // btnDealWithManyDuplicates
            // 
            this.btnDealWithManyDuplicates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDealWithManyDuplicates.Description = "Duplicates (Cols: Many)";
            this.btnDealWithManyDuplicates.Image = ((System.Drawing.Image)(resources.GetObject("btnDealWithManyDuplicates.Image")));
            this.btnDealWithManyDuplicates.Label = "Duplicates (Cols: Many)";
            this.btnDealWithManyDuplicates.Name = "btnDealWithManyDuplicates";
            this.btnDealWithManyDuplicates.ShowImage = true;
            this.btnDealWithManyDuplicates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDealWithManyDuplicates_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // btnReadUsersGroupMembership
            // 
            this.btnReadUsersGroupMembership.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadUsersGroupMembership.Description = "Active Directory User Members - Sheet Name";
            this.btnReadUsersGroupMembership.Image = ((System.Drawing.Image)(resources.GetObject("btnReadUsersGroupMembership.Image")));
            this.btnReadUsersGroupMembership.Label = "Users - SheetName";
            this.btnReadUsersGroupMembership.Name = "btnReadUsersGroupMembership";
            this.btnReadUsersGroupMembership.ShowImage = true;
            this.btnReadUsersGroupMembership.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadUsersGroupMembershipSheetName_Click);
            // 
            // btnReadUsersGroupMembershipActiveCell
            // 
            this.btnReadUsersGroupMembershipActiveCell.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadUsersGroupMembershipActiveCell.Description = "Active Directory Group Members - Active Cell";
            this.btnReadUsersGroupMembershipActiveCell.Image = ((System.Drawing.Image)(resources.GetObject("btnReadUsersGroupMembershipActiveCell.Image")));
            this.btnReadUsersGroupMembershipActiveCell.Label = "User (Groups) - Active Cell";
            this.btnReadUsersGroupMembershipActiveCell.Name = "btnReadUsersGroupMembershipActiveCell";
            this.btnReadUsersGroupMembershipActiveCell.ShowImage = true;
            this.btnReadUsersGroupMembershipActiveCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadUsersGroupMembershipActiveCell_Click);
            // 
            // btnLoadADGroupIntoSpreadsheet
            // 
            this.btnLoadADGroupIntoSpreadsheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoadADGroupIntoSpreadsheet.Description = "Active Directory Group Members - Sheet Name";
            this.btnLoadADGroupIntoSpreadsheet.Image = ((System.Drawing.Image)(resources.GetObject("btnLoadADGroupIntoSpreadsheet.Image")));
            this.btnLoadADGroupIntoSpreadsheet.Label = "Groups - SheetName";
            this.btnLoadADGroupIntoSpreadsheet.Name = "btnLoadADGroupIntoSpreadsheet";
            this.btnLoadADGroupIntoSpreadsheet.ShowImage = true;
            this.btnLoadADGroupIntoSpreadsheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadGroupUsersMembershipSheetName_Click);
            // 
            // btnLoadADGroupIntoSpreadsheetActiveCell
            // 
            this.btnLoadADGroupIntoSpreadsheetActiveCell.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoadADGroupIntoSpreadsheetActiveCell.Description = "Active Directory Group Members - Active Cell";
            this.btnLoadADGroupIntoSpreadsheetActiveCell.Image = ((System.Drawing.Image)(resources.GetObject("btnLoadADGroupIntoSpreadsheetActiveCell.Image")));
            this.btnLoadADGroupIntoSpreadsheetActiveCell.Label = "Groups (Users) - Active Cell";
            this.btnLoadADGroupIntoSpreadsheetActiveCell.Name = "btnLoadADGroupIntoSpreadsheetActiveCell";
            this.btnLoadADGroupIntoSpreadsheetActiveCell.ShowImage = true;
            this.btnLoadADGroupIntoSpreadsheetActiveCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadGroupUsersMembershipActiveCell_Click);
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // btnWriteTimeSheet
            // 
            this.btnWriteTimeSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWriteTimeSheet.Description = "Zap Worksheet";
            this.btnWriteTimeSheet.Image = ((System.Drawing.Image)(resources.GetObject("btnWriteTimeSheet.Image")));
            this.btnWriteTimeSheet.Label = "Update timesheet";
            this.btnWriteTimeSheet.Name = "btnWriteTimeSheet";
            this.btnWriteTimeSheet.ShowImage = true;
            this.btnWriteTimeSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWriteTimeSheet_Click);
            // 
            // btnPingServers
            // 
            this.btnPingServers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPingServers.Description = "Zap Worksheet";
            this.btnPingServers.Image = ((System.Drawing.Image)(resources.GetObject("btnPingServers.Image")));
            this.btnPingServers.Label = "Ping Servers";
            this.btnPingServers.Name = "btnPingServers";
            this.btnPingServers.ShowImage = true;
            this.btnPingServers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPingServers_Click);
            // 
            // ExcelRibbon
            // 
            this.Name = "ExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.MYTOOLBAR);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.MYTOOLBAR.ResumeLayout(false);
            this.MYTOOLBAR.PerformLayout();
            this.customToolbar.ResumeLayout(false);
            this.customToolbar.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab MYTOOLBAR;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup customToolbar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadFolders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompareSheets;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesA;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonDeleteLines;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDealWithSingleDuplicates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDealWithManyDuplicates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadADGroupIntoSpreadsheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadADGroupIntoSpreadsheetActiveCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadUsersGroupMembership;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadUsersGroupMembershipActiveCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWriteTimeSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPingServers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
