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
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnReadFolders = this.Factory.CreateRibbonButton();
            this.btnCompareSheets = this.Factory.CreateRibbonButton();
            this.btnZap = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLinesA = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLinesB = this.Factory.CreateRibbonButton();
            this.btnDeleteBlankLinesC = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
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
            this.customToolbar.Items.Add(this.btnZap);
            this.customToolbar.Items.Add(this.splitButton1);
            this.customToolbar.Items.Add(this.btnCompareSheets);
            this.customToolbar.Label = "Custom Toolbar";
            this.customToolbar.Name = "customToolbar";
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
            // btnDeleteBlankLinesA
            // 
            this.btnDeleteBlankLinesA.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesA.Image")));
            this.btnDeleteBlankLinesA.Label = "Mode: A";
            this.btnDeleteBlankLinesA.Name = "btnDeleteBlankLinesA";
            this.btnDeleteBlankLinesA.ShowImage = true;
            this.btnDeleteBlankLinesA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesA_Click);
            // 
            // btnDeleteBlankLinesB
            // 
            this.btnDeleteBlankLinesB.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesB.Image")));
            this.btnDeleteBlankLinesB.Label = "Mode: B";
            this.btnDeleteBlankLinesB.Name = "btnDeleteBlankLinesB";
            this.btnDeleteBlankLinesB.ScreenTip = "Fastest";
            this.btnDeleteBlankLinesB.ShowImage = true;
            this.btnDeleteBlankLinesB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesB_Click);
            // 
            // btnDeleteBlankLinesC
            // 
            this.btnDeleteBlankLinesC.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteBlankLinesC.Image")));
            this.btnDeleteBlankLinesC.Label = "Mode: C";
            this.btnDeleteBlankLinesC.Name = "btnDeleteBlankLinesC";
            this.btnDeleteBlankLinesC.ScreenTip = "Slowest";
            this.btnDeleteBlankLinesC.ShowImage = true;
            this.btnDeleteBlankLinesC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteBlankLinesC_Click);
            // 
            // splitButton1
            // 
            this.splitButton1.ButtonType = Microsoft.Office.Tools.Ribbon.RibbonButtonType.ToggleButton;
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Description = "Mode type";
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.btnDeleteBlankLinesA);
            this.splitButton1.Items.Add(this.btnDeleteBlankLinesB);
            this.splitButton1.Items.Add(this.btnDeleteBlankLinesC);
            this.splitButton1.Label = "Delete Blank Lions";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton1_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReadFolders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompareSheets;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesB;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteBlankLinesA;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
