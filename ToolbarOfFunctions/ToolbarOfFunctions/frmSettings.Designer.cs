#pragma warning disable IDE1006 // Naming Styles

namespace ToolbarOfFunctions
{
    partial class frmSettings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSettings));
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.pingServers = new System.Windows.Forms.GroupBox();
            this.numColPingRead = new System.Windows.Forms.NumericUpDown();
            this.btnCancel = new System.Windows.Forms.Button();
            this.numColPingWrite = new System.Windows.Forms.NumericUpDown();
            this.numPingSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.grpTimeSheet = new System.Windows.Forms.GroupBox();
            this.chkTimeSheetGetRowNo = new System.Windows.Forms.CheckBox();
            this.numTimeSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.grpCompare = new System.Windows.Forms.GroupBox();
            this.txtColourNotFound = new System.Windows.Forms.TextBox();
            this.txtColourFound = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.numComparingStartRow = new System.Windows.Forms.NumericUpDown();
            this.numDupliateColumnToCheck = new System.Windows.Forms.NumericUpDown();
            this.numNoOfColumnsToCheck = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.numHighlightRowsOver = new System.Windows.Forms.NumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cmboDelModeAorBorC = new System.Windows.Forms.ComboBox();
            this.chkDisplayTimeTaken = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmboHighLightOrDelete = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmboShowToolbarDescription = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmboDifferences = new System.Windows.Forms.ComboBox();
            this.chkProduceMessageBox = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.pingServers.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).BeginInit();
            this.grpTimeSheet.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).BeginInit();
            this.grpCompare.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pingServers
            // 
            this.pingServers.Controls.Add(this.numColPingRead);
            this.pingServers.Controls.Add(this.numColPingWrite);
            this.pingServers.Controls.Add(this.numPingSheetRowNo);
            this.pingServers.Controls.Add(this.label12);
            this.pingServers.Controls.Add(this.label13);
            this.pingServers.Controls.Add(this.label14);
            this.pingServers.Dock = System.Windows.Forms.DockStyle.Top;
            this.pingServers.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pingServers.ForeColor = System.Drawing.SystemColors.ControlText;
            this.pingServers.Location = new System.Drawing.Point(0, 0);
            this.pingServers.Name = "pingServers";
            this.pingServers.Size = new System.Drawing.Size(572, 120);
            this.pingServers.TabIndex = 32;
            this.pingServers.TabStop = false;
            this.pingServers.Text = "Ping Servers ..:";
            // 
            // numColPingRead
            // 
            this.numColPingRead.Location = new System.Drawing.Point(274, 48);
            this.numColPingRead.Name = "numColPingRead";
            this.numColPingRead.Size = new System.Drawing.Size(61, 23);
            this.numColPingRead.TabIndex = 37;
            this.numColPingRead.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // btnCancel
            // 
            this.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCancel.Location = new System.Drawing.Point(473, 184);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(87, 27);
            this.btnCancel.TabIndex = 33;
            this.btnCancel.Text = "Exit";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // numColPingWrite
            // 
            this.numColPingWrite.Location = new System.Drawing.Point(272, 77);
            this.numColPingWrite.Name = "numColPingWrite";
            this.numColPingWrite.Size = new System.Drawing.Size(61, 23);
            this.numColPingWrite.TabIndex = 36;
            this.numColPingWrite.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // numPingSheetRowNo
            // 
            this.numPingSheetRowNo.Location = new System.Drawing.Point(274, 19);
            this.numPingSheetRowNo.Name = "numPingSheetRowNo";
            this.numPingSheetRowNo.Size = new System.Drawing.Size(61, 23);
            this.numPingSheetRowNo.TabIndex = 35;
            this.numPingSheetRowNo.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(133, 76);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(133, 15);
            this.label12.TabIndex = 34;
            this.label12.Text = "Ping Write Column:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(140, 50);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(126, 15);
            this.label13.TabIndex = 33;
            this.label13.Text = "Ping Read Column:";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(189, 19);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(77, 15);
            this.label14.TabIndex = 32;
            this.label14.Text = "Start Row:";
            // 
            // grpTimeSheet
            // 
            this.grpTimeSheet.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.grpTimeSheet.Controls.Add(this.chkTimeSheetGetRowNo);
            this.grpTimeSheet.Controls.Add(this.numTimeSheetRowNo);
            this.grpTimeSheet.Controls.Add(this.label11);
            this.grpTimeSheet.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpTimeSheet.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpTimeSheet.Location = new System.Drawing.Point(0, 120);
            this.grpTimeSheet.Name = "grpTimeSheet";
            this.grpTimeSheet.Size = new System.Drawing.Size(572, 81);
            this.grpTimeSheet.TabIndex = 34;
            this.grpTimeSheet.TabStop = false;
            this.grpTimeSheet.Text = "Timesheet ..:";
            // 
            // chkTimeSheetGetRowNo
            // 
            this.chkTimeSheetGetRowNo.AutoSize = true;
            this.chkTimeSheetGetRowNo.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkTimeSheetGetRowNo.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTimeSheetGetRowNo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.Location = new System.Drawing.Point(147, 50);
            this.chkTimeSheetGetRowNo.Name = "chkTimeSheetGetRowNo";
            this.chkTimeSheetGetRowNo.Size = new System.Drawing.Size(135, 19);
            this.chkTimeSheetGetRowNo.TabIndex = 28;
            this.chkTimeSheetGetRowNo.Text = "Get next row No:";
            this.chkTimeSheetGetRowNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.UseVisualStyleBackColor = true;
            // 
            // numTimeSheetRowNo
            // 
            this.numTimeSheetRowNo.Location = new System.Drawing.Point(275, 21);
            this.numTimeSheetRowNo.Maximum = new decimal(new int[] {
            300000,
            0,
            0,
            0});
            this.numTimeSheetRowNo.Name = "numTimeSheetRowNo";
            this.numTimeSheetRowNo.Size = new System.Drawing.Size(61, 23);
            this.numTimeSheetRowNo.TabIndex = 27;
            this.numTimeSheetRowNo.Value = new decimal(new int[] {
            3754,
            0,
            0,
            0});
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(98, 22);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(168, 15);
            this.label11.TabIndex = 26;
            this.label11.Text = "Timesheet Row Start No:";
            // 
            // grpCompare
            // 
            this.grpCompare.Controls.Add(this.txtColourNotFound);
            this.grpCompare.Controls.Add(this.txtColourFound);
            this.grpCompare.Controls.Add(this.label9);
            this.grpCompare.Controls.Add(this.label8);
            this.grpCompare.Controls.Add(this.numComparingStartRow);
            this.grpCompare.Controls.Add(this.numDupliateColumnToCheck);
            this.grpCompare.Controls.Add(this.numNoOfColumnsToCheck);
            this.grpCompare.Controls.Add(this.label6);
            this.grpCompare.Controls.Add(this.label5);
            this.grpCompare.Controls.Add(this.label4);
            this.grpCompare.Dock = System.Windows.Forms.DockStyle.Top;
            this.grpCompare.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpCompare.Location = new System.Drawing.Point(0, 201);
            this.grpCompare.Name = "grpCompare";
            this.grpCompare.Size = new System.Drawing.Size(572, 191);
            this.grpCompare.TabIndex = 35;
            this.grpCompare.TabStop = false;
            this.grpCompare.Text = "Compare ..:";
            // 
            // txtColourNotFound
            // 
            this.txtColourNotFound.ForeColor = System.Drawing.Color.Blue;
            this.txtColourNotFound.Location = new System.Drawing.Point(277, 155);
            this.txtColourNotFound.Name = "txtColourNotFound";
            this.txtColourNotFound.Size = new System.Drawing.Size(116, 23);
            this.txtColourNotFound.TabIndex = 24;
            this.txtColourNotFound.Text = "Not Found";
            // 
            // txtColourFound
            // 
            this.txtColourFound.ForeColor = System.Drawing.Color.Red;
            this.txtColourFound.Location = new System.Drawing.Point(277, 122);
            this.txtColourFound.Name = "txtColourFound";
            this.txtColourFound.Size = new System.Drawing.Size(116, 23);
            this.txtColourFound.TabIndex = 23;
            this.txtColourFound.Text = "Found";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(42, 162);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(224, 15);
            this.label9.TabIndex = 22;
            this.label9.Text = "Colour to mark identicle items:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(42, 125);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(224, 15);
            this.label8.TabIndex = 21;
            this.label8.Text = "Colour to mark identicle items:";
            // 
            // numComparingStartRow
            // 
            this.numComparingStartRow.Location = new System.Drawing.Point(275, 56);
            this.numComparingStartRow.Name = "numComparingStartRow";
            this.numComparingStartRow.Size = new System.Drawing.Size(61, 23);
            this.numComparingStartRow.TabIndex = 19;
            this.numComparingStartRow.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // numDupliateColumnToCheck
            // 
            this.numDupliateColumnToCheck.Location = new System.Drawing.Point(275, 89);
            this.numDupliateColumnToCheck.Name = "numDupliateColumnToCheck";
            this.numDupliateColumnToCheck.Size = new System.Drawing.Size(61, 23);
            this.numDupliateColumnToCheck.TabIndex = 18;
            this.numDupliateColumnToCheck.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // numNoOfColumnsToCheck
            // 
            this.numNoOfColumnsToCheck.Location = new System.Drawing.Point(275, 23);
            this.numNoOfColumnsToCheck.Name = "numNoOfColumnsToCheck";
            this.numNoOfColumnsToCheck.Size = new System.Drawing.Size(61, 23);
            this.numNoOfColumnsToCheck.TabIndex = 17;
            this.numNoOfColumnsToCheck.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(56, 89);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(210, 15);
            this.label6.TabIndex = 16;
            this.label6.Text = "Duplicate Check Start Column:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(168, 57);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 15);
            this.label5.TabIndex = 15;
            this.label5.Text = "Starting Row:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(98, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(168, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "No Of Columns To Check:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.numHighlightRowsOver);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.cmboDelModeAorBorC);
            this.groupBox1.Controls.Add(this.chkDisplayTimeTaken);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cmboHighLightOrDelete);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.cmboShowToolbarDescription);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmboDifferences);
            this.groupBox1.Controls.Add(this.chkProduceMessageBox);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 392);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(572, 224);
            this.groupBox1.TabIndex = 36;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Misc ..:";
            // 
            // numHighlightRowsOver
            // 
            this.numHighlightRowsOver.Location = new System.Drawing.Point(275, 188);
            this.numHighlightRowsOver.Maximum = new decimal(new int[] {
            1024,
            0,
            0,
            0});
            this.numHighlightRowsOver.Name = "numHighlightRowsOver";
            this.numHighlightRowsOver.Size = new System.Drawing.Size(61, 23);
            this.numHighlightRowsOver.TabIndex = 34;
            this.numHighlightRowsOver.Value = new decimal(new int[] {
            180,
            0,
            0,
            0});
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(98, 186);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(168, 15);
            this.label10.TabIndex = 33;
            this.label10.Text = "Highlight lengths over:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(70, 156);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(196, 15);
            this.label7.TabIndex = 32;
            this.label7.Text = "Del blank Lines Mode A or B";
            // 
            // cmboDelModeAorBorC
            // 
            this.cmboDelModeAorBorC.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDelModeAorBorC.FormattingEnabled = true;
            this.cmboDelModeAorBorC.Location = new System.Drawing.Point(275, 159);
            this.cmboDelModeAorBorC.Name = "cmboDelModeAorBorC";
            this.cmboDelModeAorBorC.Size = new System.Drawing.Size(161, 23);
            this.cmboDelModeAorBorC.TabIndex = 31;
            // 
            // chkDisplayTimeTaken
            // 
            this.chkDisplayTimeTaken.AutoSize = true;
            this.chkDisplayTimeTaken.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkDisplayTimeTaken.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDisplayTimeTaken.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Location = new System.Drawing.Point(126, 109);
            this.chkDisplayTimeTaken.Name = "chkDisplayTimeTaken";
            this.chkDisplayTimeTaken.Size = new System.Drawing.Size(156, 19);
            this.chkDisplayTimeTaken.TabIndex = 30;
            this.chkDisplayTimeTaken.Text = "Display Time Taken:";
            this.chkDisplayTimeTaken.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(252, 15);
            this.label3.TabIndex = 29;
            this.label3.Text = "Highlight or Delete Duplicate rows:";
            // 
            // cmboHighLightOrDelete
            // 
            this.cmboHighLightOrDelete.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboHighLightOrDelete.FormattingEnabled = true;
            this.cmboHighLightOrDelete.Location = new System.Drawing.Point(274, 80);
            this.cmboHighLightOrDelete.Name = "cmboHighLightOrDelete";
            this.cmboHighLightOrDelete.Size = new System.Drawing.Size(161, 23);
            this.cmboHighLightOrDelete.TabIndex = 28;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(72, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(196, 15);
            this.label2.TabIndex = 27;
            this.label2.Text = "Show Tool Bar Descriptions:";
            // 
            // cmboShowToolbarDescription
            // 
            this.cmboShowToolbarDescription.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboShowToolbarDescription.FormattingEnabled = true;
            this.cmboShowToolbarDescription.Location = new System.Drawing.Point(274, 22);
            this.cmboShowToolbarDescription.Name = "cmboShowToolbarDescription";
            this.cmboShowToolbarDescription.Size = new System.Drawing.Size(161, 23);
            this.cmboShowToolbarDescription.TabIndex = 26;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(177, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 15);
            this.label1.TabIndex = 25;
            this.label1.Text = "Differences:";
            // 
            // cmboDifferences
            // 
            this.cmboDifferences.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDifferences.FormattingEnabled = true;
            this.cmboDifferences.Location = new System.Drawing.Point(274, 51);
            this.cmboDifferences.Name = "cmboDifferences";
            this.cmboDifferences.Size = new System.Drawing.Size(161, 23);
            this.cmboDifferences.TabIndex = 24;
            // 
            // chkProduceMessageBox
            // 
            this.chkProduceMessageBox.AutoSize = true;
            this.chkProduceMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceMessageBox.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.Location = new System.Drawing.Point(119, 134);
            this.chkProduceMessageBox.Name = "chkProduceMessageBox";
            this.chkProduceMessageBox.Size = new System.Drawing.Size(163, 19);
            this.chkProduceMessageBox.TabIndex = 23;
            this.chkProduceMessageBox.Text = "Produce Message Box?";
            this.chkProduceMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripDropDownButton1,
            this.toolStripStatusLabel2,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 622);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(572, 22);
            this.statusStrip1.TabIndex = 37;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(29, 20);
            this.toolStripDropDownButton1.Text = "toolStripDropDownButton1";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel2.Text = "toolStripStatusLabel2";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(572, 644);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpCompare);
            this.Controls.Add(this.grpTimeSheet);
            this.Controls.Add(this.pingServers);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmSettings";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings ...";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.pingServers.ResumeLayout(false);
            this.pingServers.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).EndInit();
            this.grpTimeSheet.ResumeLayout(false);
            this.grpTimeSheet.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).EndInit();
            this.grpCompare.ResumeLayout(false);
            this.grpCompare.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.GroupBox pingServers;
        private System.Windows.Forms.NumericUpDown numColPingRead;
        private System.Windows.Forms.NumericUpDown numColPingWrite;
        private System.Windows.Forms.NumericUpDown numPingSheetRowNo;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox grpTimeSheet;
        public System.Windows.Forms.CheckBox chkTimeSheetGetRowNo;
        private System.Windows.Forms.NumericUpDown numTimeSheetRowNo;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.GroupBox grpCompare;
        private System.Windows.Forms.TextBox txtColourNotFound;
        private System.Windows.Forms.TextBox txtColourFound;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numComparingStartRow;
        private System.Windows.Forms.NumericUpDown numDupliateColumnToCheck;
        private System.Windows.Forms.NumericUpDown numNoOfColumnsToCheck;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.NumericUpDown numHighlightRowsOver;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmboDelModeAorBorC;
        public System.Windows.Forms.CheckBox chkDisplayTimeTaken;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmboHighLightOrDelete;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmboShowToolbarDescription;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmboDifferences;
        public System.Windows.Forms.CheckBox chkProduceMessageBox;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}