namespace ToolbarOfFunctions
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.MYTOOLBAR = this.Factory.CreateRibbonTab();
            this.customToolbar = this.Factory.CreateRibbonGroup();
            this.btnZap = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLines = this.Factory.CreateRibbonButton();
            this.btnReadFolders = this.Factory.CreateRibbonButton();
            this.btnCompareSheets = this.Factory.CreateRibbonButton();
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
            this.customToolbar.Items.Add(this.btnReadFolders);
            this.customToolbar.Items.Add(this.btnCompareSheets);
            this.customToolbar.Items.Add(this.btnZap);
            this.customToolbar.Items.Add(this.btnDeleteBlankLines);
            this.customToolbar.Label = "Custom Toolbar";
            this.customToolbar.Name = "customToolbar";
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
            // btnDeleteBlankLines
            // 
            this.btnDeleteBlankLines.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteBlankLines.Description = "Delete Blank Lines";
            this.btnDeleteBlankLines.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLines.Image")));
            this.btnDeleteBlankLines.Label = "Delete Blank Lines";
            this.btnDeleteBlankLines.Name = "btnDeleteBlankLines";
            this.btnDeleteBlankLines.ShowImage = true;
            this.btnDeleteBlankLines.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLines_Click);
            // 
            // btnReadFolders
            // 
            this.btnReadFolders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReadFolders.Image = ((System.Drawing.Image)(resources.GetObject("btnReadFolders.Image")));
            this.btnReadFolders.Label = "Read Folders";
            this.btnReadFolders.Name = "btnReadFolders";
            this.btnReadFolders.ShowImage = true;
            this.btnReadFolders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReadFolders_Click);
            // 
            // btnCompareSheets
            // 
            this.btnCompareSheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCompareSheets.Image = ((System.Drawing.Image)(resources.GetObject("btnCompareSheets.Image")));
            this.btnCompareSheets.Label = "Compare Sheets";
            this.btnCompareSheets.Name = "btnCompareSheets";
            this.btnCompareSheets.ShowImage = true;
            this.btnCompareSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompareSheets_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLines;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadFolders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompareSheets;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
