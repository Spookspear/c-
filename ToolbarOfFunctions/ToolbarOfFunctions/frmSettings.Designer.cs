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
            this.btnCancel = new System.Windows.Forms.Button();
            this.numPingSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.chkTimeSheetGetRowNo = new System.Windows.Forms.CheckBox();
            this.numTimeSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
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
            this.chkHideText = new System.Windows.Forms.CheckBox();
            this.chkLargeButtons = new System.Windows.Forms.CheckBox();
            this.numHighlightRowsOver = new System.Windows.Forms.NumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.cmboDelModeAorBorC = new System.Windows.Forms.ComboBox();
            this.chkDisplayTimeTaken = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmboHighLightOrDelete = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmboCompareDifferences = new System.Windows.Forms.ComboBox();
            this.chkProduceMessageBox = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnApply
            // 
            this.btnApply.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnApply.Location = new System.Drawing.Point(982, 684);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(87, 27);
            this.btnApply.TabIndex = 38;
            this.btnApply.Text = "&Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // numColPingRead
            // 
            this.numColPingRead.Location = new System.Drawing.Point(593, 559);
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
            this.numColPingWrite.Location = new System.Drawing.Point(593, 588);
            this.numColPingWrite.Name = "numColPingWrite";
            this.numColPingWrite.Size = new System.Drawing.Size(61, 23);
            this.numColPingWrite.TabIndex = 36;
            this.numColPingWrite.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnCancel
            // 
            this.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(1104, 684);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(87, 27);
            this.btnCancel.TabIndex = 33;
            this.btnCancel.Text = "E&xit";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // numPingSheetRowNo
            // 
            this.numPingSheetRowNo.Location = new System.Drawing.Point(593, 532);
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
            this.label12.Location = new System.Drawing.Point(454, 590);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(133, 15);
            this.label12.TabIndex = 34;
            this.label12.Text = "Ping Write Column:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(463, 561);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(126, 15);
            this.label13.TabIndex = 33;
            this.label13.Text = "Ping Read Column:";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(510, 534);
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
            this.chkTimeSheetGetRowNo.Location = new System.Drawing.Point(489, 410);
            this.chkTimeSheetGetRowNo.Name = "chkTimeSheetGetRowNo";
            this.chkTimeSheetGetRowNo.Size = new System.Drawing.Size(135, 19);
            this.chkTimeSheetGetRowNo.TabIndex = 28;
            this.chkTimeSheetGetRowNo.Text = "Get next row No:";
            this.chkTimeSheetGetRowNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.UseVisualStyleBackColor = true;
            // 
            // numTimeSheetRowNo
            // 
            this.numTimeSheetRowNo.Location = new System.Drawing.Point(614, 381);
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
            this.label11.Location = new System.Drawing.Point(437, 383);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(168, 15);
            this.label11.TabIndex = 26;
            this.label11.Text = "Timesheet Row Start No:";
            // 
            // txtColourNotFound
            // 
            this.txtColourNotFound.ForeColor = System.Drawing.Color.Blue;
            this.txtColourNotFound.Location = new System.Drawing.Point(999, 194);
            this.txtColourNotFound.Name = "txtColourNotFound";
            this.txtColourNotFound.Size = new System.Drawing.Size(61, 23);
            this.txtColourNotFound.TabIndex = 24;
            this.txtColourNotFound.Text = "Not Found";
            // 
            // txtColourFound
            // 
            this.txtColourFound.ForeColor = System.Drawing.Color.Red;
            this.txtColourFound.Location = new System.Drawing.Point(1001, 168);
            this.txtColourFound.Name = "txtColourFound";
            this.txtColourFound.Size = new System.Drawing.Size(59, 23);
            this.txtColourFound.TabIndex = 23;
            this.txtColourFound.Text = "Found";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(769, 202);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(224, 15);
            this.label9.TabIndex = 22;
            this.label9.Text = "Colour to mark identicle items:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(766, 171);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(224, 15);
            this.label8.TabIndex = 21;
            this.label8.Text = "Colour to mark identicle items:";
            // 
            // numComparingStartRow
            // 
            this.numComparingStartRow.Location = new System.Drawing.Point(999, 110);
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
            this.numDupliateColumnToCheck.Location = new System.Drawing.Point(998, 139);
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
            this.numNoOfColumnsToCheck.Location = new System.Drawing.Point(999, 81);
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
            this.label6.Location = new System.Drawing.Point(782, 141);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(210, 15);
            this.label6.TabIndex = 16;
            this.label6.Text = "Duplicate Check Start Column:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(892, 112);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 15);
            this.label5.TabIndex = 15;
            this.label5.Text = "Starting Row:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(822, 83);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(168, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "No Of Columns To Check:";
            // 
            // chkHideText
            // 
            this.chkHideText.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkHideText.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHideText.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.Location = new System.Drawing.Point(457, 58);
            this.chkHideText.Name = "chkHideText";
            this.chkHideText.Size = new System.Drawing.Size(142, 19);
            this.chkHideText.TabIndex = 36;
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
            this.chkLargeButtons.Location = new System.Drawing.Point(478, 33);
            this.chkLargeButtons.Name = "chkLargeButtons";
            this.chkLargeButtons.Size = new System.Drawing.Size(107, 17);
            this.chkLargeButtons.TabIndex = 35;
            this.chkLargeButtons.Text = "Large Buttons:";
            this.chkLargeButtons.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkLargeButtons.UseVisualStyleBackColor = true;
            // 
            // numHighlightRowsOver
            // 
            this.numHighlightRowsOver.Location = new System.Drawing.Point(567, 248);
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
            this.label10.Location = new System.Drawing.Point(411, 248);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(145, 13);
            this.label10.TabIndex = 33;
            this.label10.Text = "Highlight lengths over:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(374, 212);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(187, 13);
            this.label7.TabIndex = 32;
            this.label7.Text = "Del blank Lines Mode A, B or C";
            // 
            // cmboDelModeAorBorC
            // 
            this.cmboDelModeAorBorC.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDelModeAorBorC.FormattingEnabled = true;
            this.cmboDelModeAorBorC.Location = new System.Drawing.Point(567, 202);
            this.cmboDelModeAorBorC.Name = "cmboDelModeAorBorC";
            this.cmboDelModeAorBorC.Size = new System.Drawing.Size(161, 23);
            this.cmboDelModeAorBorC.TabIndex = 31;
            // 
            // chkDisplayTimeTaken
            // 
            this.chkDisplayTimeTaken.AutoSize = true;
            this.chkDisplayTimeTaken.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkDisplayTimeTaken.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDisplayTimeTaken.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Location = new System.Drawing.Point(443, 156);
            this.chkDisplayTimeTaken.Name = "chkDisplayTimeTaken";
            this.chkDisplayTimeTaken.Size = new System.Drawing.Size(137, 17);
            this.chkDisplayTimeTaken.TabIndex = 30;
            this.chkDisplayTimeTaken.Text = "Display Time Taken:";
            this.chkDisplayTimeTaken.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(363, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(217, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "Highlight or Delete Duplicate rows:";
            // 
            // cmboHighLightOrDelete
            // 
            this.cmboHighLightOrDelete.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboHighLightOrDelete.FormattingEnabled = true;
            this.cmboHighLightOrDelete.Location = new System.Drawing.Point(588, 126);
            this.cmboHighLightOrDelete.Name = "cmboHighLightOrDelete";
            this.cmboHighLightOrDelete.Size = new System.Drawing.Size(161, 23);
            this.cmboHighLightOrDelete.TabIndex = 28;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(450, 102);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Compare Differences:";
            // 
            // cmboCompareDifferences
            // 
            this.cmboCompareDifferences.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboCompareDifferences.FormattingEnabled = true;
            this.cmboCompareDifferences.Location = new System.Drawing.Point(588, 97);
            this.cmboCompareDifferences.Name = "cmboCompareDifferences";
            this.cmboCompareDifferences.Size = new System.Drawing.Size(161, 23);
            this.cmboCompareDifferences.TabIndex = 24;
            // 
            // chkProduceMessageBox
            // 
            this.chkProduceMessageBox.AutoSize = true;
            this.chkProduceMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceMessageBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.Location = new System.Drawing.Point(432, 179);
            this.chkProduceMessageBox.Name = "chkProduceMessageBox";
            this.chkProduceMessageBox.Size = new System.Drawing.Size(143, 17);
            this.chkProduceMessageBox.TabIndex = 23;
            this.chkProduceMessageBox.Text = "Produce Message Box?";
            this.chkProduceMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(342, 532);
            this.tabControl1.TabIndex = 39;
            // 
            // tabPage1
            // 
            this.tabPage1.Location = new System.Drawing.Point(26, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(312, 524);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(26, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(312, 524);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1203, 723);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.chkTimeSheetGetRowNo);
            this.Controls.Add(this.numColPingRead);
            this.Controls.Add(this.numColPingWrite);
            this.Controls.Add(this.txtColourNotFound);
            this.Controls.Add(this.numTimeSheetRowNo);
            this.Controls.Add(this.numPingSheetRowNo);
            this.Controls.Add(this.chkHideText);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.txtColourFound);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.chkLargeButtons);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.numComparingStartRow);
            this.Controls.Add(this.cmboDelModeAorBorC);
            this.Controls.Add(this.numDupliateColumnToCheck);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.numNoOfColumnsToCheck);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.chkDisplayTimeTaken);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.numHighlightRowsOver);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmboHighLightOrDelete);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.chkProduceMessageBox);
            this.Controls.Add(this.cmboCompareDifferences);
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
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).EndInit();
            this.tabControl1.ResumeLayout(false);
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
        private System.Windows.Forms.Button btnCancel;
        public System.Windows.Forms.CheckBox chkTimeSheetGetRowNo;
        private System.Windows.Forms.NumericUpDown numTimeSheetRowNo;
        private System.Windows.Forms.Label label11;
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
        private System.Windows.Forms.NumericUpDown numHighlightRowsOver;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmboDelModeAorBorC;
        public System.Windows.Forms.CheckBox chkDisplayTimeTaken;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.ComboBox cmboHighLightOrDelete;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox cmboCompareDifferences;
        public System.Windows.Forms.CheckBox chkProduceMessageBox;
        public System.Windows.Forms.CheckBox chkLargeButtons;
        public System.Windows.Forms.CheckBox chkHideText;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
    }
}