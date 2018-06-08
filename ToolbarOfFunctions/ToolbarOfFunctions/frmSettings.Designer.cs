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
            this.btnApply = new System.Windows.Forms.Button();
            this.numColPingRead = new System.Windows.Forms.NumericUpDown();
            this.numColPingWrite = new System.Windows.Forms.NumericUpDown();
            this.numPingSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.chkTimeSheetGetRowNo = new System.Windows.Forms.CheckBox();
            this.numTimeSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.numDupliateColumnToCheck = new System.Windows.Forms.NumericUpDown();
            this.numNoOfColumnsToCheck = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.numHighlightRowsOver = new System.Windows.Forms.NumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cmboDelModeAorBorC = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmboHighLightOrDelete = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkHideText = new System.Windows.Forms.CheckBox();
            this.chkLargeButtons = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkDisplayTimeTaken = new System.Windows.Forms.CheckBox();
            this.chkProduceCompleteMessageBox = new System.Windows.Forms.CheckBox();
            this.chkProduceInitialMessageBox = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnColourFound = new System.Windows.Forms.Button();
            this.btnColourNotFound = new System.Windows.Forms.Button();
            this.txtColourNotFound = new System.Windows.Forms.TextBox();
            this.txtColourFound = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.numComparingStartRow = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cmboCompareDifferences = new System.Windows.Forms.ComboBox();
            this.btnColourFoundBack = new System.Windows.Forms.Button();
            this.btnColourNotFoundBack = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).BeginInit();
            this.SuspendLayout();
            // 
            // btnApply
            // 
            this.btnApply.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnApply.Location = new System.Drawing.Point(337, 409);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(117, 27);
            this.btnApply.TabIndex = 38;
            this.btnApply.Text = "&Apply / Close";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // numColPingRead
            // 
            this.numColPingRead.Location = new System.Drawing.Point(219, 378);
            this.numColPingRead.Name = "numColPingRead";
            this.numColPingRead.Size = new System.Drawing.Size(61, 23);
            this.numColPingRead.TabIndex = 37;
            this.numColPingRead.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // numColPingWrite
            // 
            this.numColPingWrite.Location = new System.Drawing.Point(219, 407);
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
            this.numPingSheetRowNo.Location = new System.Drawing.Point(219, 351);
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
            this.label12.Location = new System.Drawing.Point(80, 409);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(133, 15);
            this.label12.TabIndex = 34;
            this.label12.Text = "Ping Write Column:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(89, 380);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(126, 15);
            this.label13.TabIndex = 33;
            this.label13.Text = "Ping Read Column:";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(136, 353);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(77, 15);
            this.label14.TabIndex = 32;
            this.label14.Text = "Start Row:";
            // 
            // chkTimeSheetGetRowNo
            // 
            this.chkTimeSheetGetRowNo.AutoSize = true;
            this.chkTimeSheetGetRowNo.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkTimeSheetGetRowNo.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTimeSheetGetRowNo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.Location = new System.Drawing.Point(94, 321);
            this.chkTimeSheetGetRowNo.Name = "chkTimeSheetGetRowNo";
            this.chkTimeSheetGetRowNo.Size = new System.Drawing.Size(135, 19);
            this.chkTimeSheetGetRowNo.TabIndex = 28;
            this.chkTimeSheetGetRowNo.Text = "Get next row No:";
            this.chkTimeSheetGetRowNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.UseVisualStyleBackColor = true;
            // 
            // numTimeSheetRowNo
            // 
            this.numTimeSheetRowNo.Location = new System.Drawing.Point(219, 292);
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
            this.label11.Location = new System.Drawing.Point(42, 294);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(168, 15);
            this.label11.TabIndex = 26;
            this.label11.Text = "Timesheet Row Start No:";
            // 
            // numDupliateColumnToCheck
            // 
            this.numDupliateColumnToCheck.Location = new System.Drawing.Point(528, 327);
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
            this.numNoOfColumnsToCheck.Location = new System.Drawing.Point(528, 263);
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
            this.label6.Location = new System.Drawing.Point(312, 329);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(210, 15);
            this.label6.TabIndex = 16;
            this.label6.Text = "Duplicate Check Start Column:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(351, 265);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(168, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "No Of Columns To Check:";
            // 
            // numHighlightRowsOver
            // 
            this.numHighlightRowsOver.Location = new System.Drawing.Point(219, 263);
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
            this.label10.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(63, 263);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(145, 13);
            this.label10.TabIndex = 33;
            this.label10.Text = "Highlight lengths over:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(26, 227);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(187, 13);
            this.label7.TabIndex = 32;
            this.label7.Text = "Del blank Lines Mode A, B or C";
            // 
            // cmboDelModeAorBorC
            // 
            this.cmboDelModeAorBorC.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDelModeAorBorC.FormattingEnabled = true;
            this.cmboDelModeAorBorC.Location = new System.Drawing.Point(219, 217);
            this.cmboDelModeAorBorC.Name = "cmboDelModeAorBorC";
            this.cmboDelModeAorBorC.Size = new System.Drawing.Size(161, 23);
            this.cmboDelModeAorBorC.TabIndex = 31;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(303, 303);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(217, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "Highlight or Delete Duplicate rows:";
            // 
            // cmboHighLightOrDelete
            // 
            this.cmboHighLightOrDelete.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboHighLightOrDelete.FormattingEnabled = true;
            this.cmboHighLightOrDelete.Location = new System.Drawing.Point(528, 298);
            this.cmboHighLightOrDelete.Name = "cmboHighLightOrDelete";
            this.cmboHighLightOrDelete.Size = new System.Drawing.Size(161, 23);
            this.cmboHighLightOrDelete.TabIndex = 28;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkHideText);
            this.groupBox1.Controls.Add(this.chkLargeButtons);
            this.groupBox1.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(315, 15);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(180, 90);
            this.groupBox1.TabIndex = 40;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Toolbar";
            // 
            // chkHideText
            // 
            this.chkHideText.AutoSize = true;
            this.chkHideText.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkHideText.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHideText.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.Location = new System.Drawing.Point(22, 45);
            this.chkHideText.Name = "chkHideText";
            this.chkHideText.Size = new System.Drawing.Size(125, 17);
            this.chkHideText.TabIndex = 38;
            this.chkHideText.Text = "Hide Button Text:";
            this.chkHideText.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.UseVisualStyleBackColor = true;
            // 
            // chkLargeButtons
            // 
            this.chkLargeButtons.AutoSize = true;
            this.chkLargeButtons.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkLargeButtons.Checked = true;
            this.chkLargeButtons.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLargeButtons.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkLargeButtons.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLargeButtons.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkLargeButtons.Location = new System.Drawing.Point(40, 22);
            this.chkLargeButtons.Name = "chkLargeButtons";
            this.chkLargeButtons.Size = new System.Drawing.Size(107, 17);
            this.chkLargeButtons.TabIndex = 37;
            this.chkLargeButtons.Text = "Large Buttons:";
            this.chkLargeButtons.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkLargeButtons.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkDisplayTimeTaken);
            this.groupBox2.Controls.Add(this.chkProduceCompleteMessageBox);
            this.groupBox2.Controls.Add(this.chkProduceInitialMessageBox);
            this.groupBox2.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(295, 93);
            this.groupBox2.TabIndex = 41;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Messages and Warnings";
            // 
            // chkDisplayTimeTaken
            // 
            this.chkDisplayTimeTaken.AutoSize = true;
            this.chkDisplayTimeTaken.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Checked = true;
            this.chkDisplayTimeTaken.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDisplayTimeTaken.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkDisplayTimeTaken.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDisplayTimeTaken.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Location = new System.Drawing.Point(138, 68);
            this.chkDisplayTimeTaken.Name = "chkDisplayTimeTaken";
            this.chkDisplayTimeTaken.Size = new System.Drawing.Size(137, 17);
            this.chkDisplayTimeTaken.TabIndex = 42;
            this.chkDisplayTimeTaken.Text = "Display Time Taken:";
            this.chkDisplayTimeTaken.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.UseVisualStyleBackColor = true;
            // 
            // chkProduceCompleteMessageBox
            // 
            this.chkProduceCompleteMessageBox.AutoSize = true;
            this.chkProduceCompleteMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceCompleteMessageBox.Checked = true;
            this.chkProduceCompleteMessageBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProduceCompleteMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceCompleteMessageBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceCompleteMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceCompleteMessageBox.Location = new System.Drawing.Point(48, 45);
            this.chkProduceCompleteMessageBox.Name = "chkProduceCompleteMessageBox";
            this.chkProduceCompleteMessageBox.Size = new System.Drawing.Size(227, 17);
            this.chkProduceCompleteMessageBox.TabIndex = 41;
            this.chkProduceCompleteMessageBox.Text = "Display Process Completed Message?";
            this.chkProduceCompleteMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceCompleteMessageBox.UseVisualStyleBackColor = true;
            this.chkProduceCompleteMessageBox.CheckedChanged += new System.EventHandler(this.chkProduceCompleteMessageBox_CheckedChanged);
            // 
            // chkProduceInitialMessageBox
            // 
            this.chkProduceInitialMessageBox.AutoSize = true;
            this.chkProduceInitialMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.Checked = true;
            this.chkProduceInitialMessageBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProduceInitialMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceInitialMessageBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceInitialMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.Location = new System.Drawing.Point(12, 22);
            this.chkProduceInitialMessageBox.Name = "chkProduceInitialMessageBox";
            this.chkProduceInitialMessageBox.Size = new System.Drawing.Size(263, 17);
            this.chkProduceInitialMessageBox.TabIndex = 40;
            this.chkProduceInitialMessageBox.Text = "Display Inital Confirmation Message Box?";
            this.chkProduceInitialMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnColourNotFoundBack);
            this.groupBox3.Controls.Add(this.btnColourFoundBack);
            this.groupBox3.Controls.Add(this.btnColourFound);
            this.groupBox3.Controls.Add(this.btnColourNotFound);
            this.groupBox3.Controls.Add(this.txtColourNotFound);
            this.groupBox3.Controls.Add(this.txtColourFound);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.numComparingStartRow);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.cmboCompareDifferences);
            this.groupBox3.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(12, 111);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(483, 83);
            this.groupBox3.TabIndex = 42;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Compare Sheets Options";
            // 
            // btnColourFound
            // 
            this.btnColourFound.Location = new System.Drawing.Point(342, 20);
            this.btnColourFound.Name = "btnColourFound";
            this.btnColourFound.Size = new System.Drawing.Size(47, 22);
            this.btnColourFound.TabIndex = 35;
            this.btnColourFound.Text = "Fore";
            this.btnColourFound.UseVisualStyleBackColor = true;
            this.btnColourFound.Click += new System.EventHandler(this.btnColourFound_Click);
            // 
            // btnColourNotFound
            // 
            this.btnColourNotFound.Location = new System.Drawing.Point(343, 47);
            this.btnColourNotFound.Name = "btnColourNotFound";
            this.btnColourNotFound.Size = new System.Drawing.Size(46, 22);
            this.btnColourNotFound.TabIndex = 34;
            this.btnColourNotFound.Text = "Fore";
            this.btnColourNotFound.UseVisualStyleBackColor = true;
            this.btnColourNotFound.Click += new System.EventHandler(this.btnColourNotFound_Click);
            // 
            // txtColourNotFound
            // 
            this.txtColourNotFound.ForeColor = System.Drawing.Color.Blue;
            this.txtColourNotFound.Location = new System.Drawing.Point(282, 49);
            this.txtColourNotFound.Name = "txtColourNotFound";
            this.txtColourNotFound.Size = new System.Drawing.Size(59, 20);
            this.txtColourNotFound.TabIndex = 33;
            this.txtColourNotFound.Text = "Not Found";
            this.txtColourNotFound.Click += new System.EventHandler(this.btnColourNotFound_Click);
            // 
            // txtColourFound
            // 
            this.txtColourFound.ForeColor = System.Drawing.Color.Red;
            this.txtColourFound.Location = new System.Drawing.Point(281, 19);
            this.txtColourFound.Name = "txtColourFound";
            this.txtColourFound.Size = new System.Drawing.Size(59, 20);
            this.txtColourFound.TabIndex = 32;
            this.txtColourFound.Text = "Found";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(173, 52);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(103, 13);
            this.label9.TabIndex = 31;
            this.label9.Text = "Not Found items:";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.label8.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(196, 23);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(79, 13);
            this.label8.TabIndex = 30;
            this.label8.Text = "Found items:";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // numComparingStartRow
            // 
            this.numComparingStartRow.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numComparingStartRow.Location = new System.Drawing.Point(103, 50);
            this.numComparingStartRow.Name = "numComparingStartRow";
            this.numComparingStartRow.Size = new System.Drawing.Size(64, 20);
            this.numComparingStartRow.TabIndex = 29;
            this.numComparingStartRow.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(6, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 13);
            this.label5.TabIndex = 28;
            this.label5.Text = "Starting Row:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(18, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 27;
            this.label1.Text = "Differences:";
            // 
            // cmboCompareDifferences
            // 
            this.cmboCompareDifferences.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboCompareDifferences.FormattingEnabled = true;
            this.cmboCompareDifferences.Location = new System.Drawing.Point(103, 20);
            this.cmboCompareDifferences.Name = "cmboCompareDifferences";
            this.cmboCompareDifferences.Size = new System.Drawing.Size(64, 21);
            this.cmboCompareDifferences.TabIndex = 26;
            this.cmboCompareDifferences.Text = "Colour";
            this.cmboCompareDifferences.SelectedIndexChanged += new System.EventHandler(this.cmboCompareDifferences_SelectedIndexChanged);
            // 
            // btnColourFoundBack
            // 
            this.btnColourFoundBack.Location = new System.Drawing.Point(395, 20);
            this.btnColourFoundBack.Name = "btnColourFoundBack";
            this.btnColourFoundBack.Size = new System.Drawing.Size(47, 22);
            this.btnColourFoundBack.TabIndex = 36;
            this.btnColourFoundBack.Text = "Back";
            this.btnColourFoundBack.UseVisualStyleBackColor = true;
            this.btnColourFoundBack.Click += new System.EventHandler(this.btnColourFoundBack_Click);
            // 
            // btnColourNotFoundBack
            // 
            this.btnColourNotFoundBack.Location = new System.Drawing.Point(395, 47);
            this.btnColourNotFoundBack.Name = "btnColourNotFoundBack";
            this.btnColourNotFoundBack.Size = new System.Drawing.Size(47, 22);
            this.btnColourNotFoundBack.TabIndex = 37;
            this.btnColourNotFoundBack.Text = "Back";
            this.btnColourNotFoundBack.UseVisualStyleBackColor = true;
            this.btnColourNotFoundBack.Click += new System.EventHandler(this.btnColourNotFoundBack_Click);
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(715, 474);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.chkTimeSheetGetRowNo);
            this.Controls.Add(this.numColPingRead);
            this.Controls.Add(this.numColPingWrite);
            this.Controls.Add(this.numTimeSheetRowNo);
            this.Controls.Add(this.numPingSheetRowNo);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmboDelModeAorBorC);
            this.Controls.Add(this.numDupliateColumnToCheck);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.numNoOfColumnsToCheck);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.numHighlightRowsOver);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmboHighLightOrDelete);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmSettings";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings ...";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.NumericUpDown numColPingRead;
        private System.Windows.Forms.NumericUpDown numColPingWrite;
        private System.Windows.Forms.NumericUpDown numPingSheetRowNo;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        public System.Windows.Forms.CheckBox chkTimeSheetGetRowNo;
        private System.Windows.Forms.NumericUpDown numTimeSheetRowNo;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.NumericUpDown numDupliateColumnToCheck;
        private System.Windows.Forms.NumericUpDown numNoOfColumnsToCheck;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numHighlightRowsOver;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmboDelModeAorBorC;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.ComboBox cmboHighLightOrDelete;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.CheckBox chkHideText;
        public System.Windows.Forms.CheckBox chkLargeButtons;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.CheckBox chkProduceCompleteMessageBox;
        public System.Windows.Forms.CheckBox chkProduceInitialMessageBox;
        public System.Windows.Forms.CheckBox chkDisplayTimeTaken;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox txtColourNotFound;
        private System.Windows.Forms.TextBox txtColourFound;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numComparingStartRow;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox cmboCompareDifferences;
        private System.Windows.Forms.Button btnColourNotFound;
        private System.Windows.Forms.Button btnColourFound;
        private System.Windows.Forms.Button btnColourNotFoundBack;
        private System.Windows.Forms.Button btnColourFoundBack;
    }
}