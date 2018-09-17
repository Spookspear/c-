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
            System.Windows.Forms.GroupBox groupBox6;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSettings));
            this.numColPingWrite = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.numColPingRead = new System.Windows.Forms.NumericUpDown();
            this.label13 = new System.Windows.Forms.Label();
            this.numPingSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label14 = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkHideSeperator = new System.Windows.Forms.CheckBox();
            this.chkHideText = new System.Windows.Forms.CheckBox();
            this.chkLargeButtons = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkClearFormatting = new System.Windows.Forms.CheckBox();
            this.chkTurnOffScreenValidation = new System.Windows.Forms.CheckBox();
            this.chkDisplayTimeTaken = new System.Windows.Forms.CheckBox();
            this.chkProduceCompleteMessageBox = new System.Windows.Forms.CheckBox();
            this.chkProduceInitialMessageBox = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkNotFoundBold = new System.Windows.Forms.CheckBox();
            this.chkFoundBold = new System.Windows.Forms.CheckBox();
            this.chkLBColorOrCompare = new System.Windows.Forms.CheckedListBox();
            this.btnColourNotFoundBack = new System.Windows.Forms.Button();
            this.btnColourFoundBack = new System.Windows.Forms.Button();
            this.btnColourFound = new System.Windows.Forms.Button();
            this.btnColourNotFound = new System.Windows.Forms.Button();
            this.txtColourNotFound = new System.Windows.Forms.TextBox();
            this.txtColourFound = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.numComparingStartRow = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.numDupliateColumnToCheck = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.numNoOfColumnsToCheck = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.chkLBColourOrDelete = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.numTimeSheetRowNo = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.chkTimeSheetGetRowNo = new System.Windows.Forms.CheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.numZapSheetStartRow = new System.Windows.Forms.NumericUpDown();
            this.label16 = new System.Windows.Forms.Label();
            this.numColNoForExtractedName = new System.Windows.Forms.NumericUpDown();
            this.label15 = new System.Windows.Forms.Label();
            this.chkExtractFileName = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmboWhichDate = new System.Windows.Forms.ComboBox();
            this.chkTestCode = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmboDelMode = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.numHighlightRowsOver = new System.Windows.Forms.NumericUpDown();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnTestCode = new System.Windows.Forms.Button();
            groupBox6 = new System.Windows.Forms.GroupBox();
            groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).BeginInit();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).BeginInit();
            this.groupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numZapSheetStartRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColNoForExtractedName)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).BeginInit();
            this.groupBox8.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(this.numColPingWrite);
            groupBox6.Controls.Add(this.label12);
            groupBox6.Controls.Add(this.numColPingRead);
            groupBox6.Controls.Add(this.label13);
            groupBox6.Controls.Add(this.numPingSheetRowNo);
            groupBox6.Controls.Add(this.label14);
            groupBox6.Location = new System.Drawing.Point(5, 376);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new System.Drawing.Size(452, 81);
            groupBox6.TabIndex = 49;
            groupBox6.TabStop = false;
            groupBox6.Text = "Ping Servers";
            // 
            // numColPingWrite
            // 
            this.numColPingWrite.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numColPingWrite.Location = new System.Drawing.Point(133, 47);
            this.numColPingWrite.Name = "numColPingWrite";
            this.numColPingWrite.Size = new System.Drawing.Size(52, 20);
            this.numColPingWrite.TabIndex = 41;
            this.numColPingWrite.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(14, 48);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(115, 13);
            this.label12.TabIndex = 40;
            this.label12.Text = "Ping Write Column:";
            // 
            // numColPingRead
            // 
            this.numColPingRead.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numColPingRead.Location = new System.Drawing.Point(133, 21);
            this.numColPingRead.Name = "numColPingRead";
            this.numColPingRead.Size = new System.Drawing.Size(52, 20);
            this.numColPingRead.TabIndex = 39;
            this.numColPingRead.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label13.Location = new System.Drawing.Point(21, 22);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(109, 13);
            this.label13.TabIndex = 38;
            this.label13.Text = "Ping Read Column:";
            // 
            // numPingSheetRowNo
            // 
            this.numPingSheetRowNo.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numPingSheetRowNo.Location = new System.Drawing.Point(281, 25);
            this.numPingSheetRowNo.Name = "numPingSheetRowNo";
            this.numPingSheetRowNo.Size = new System.Drawing.Size(52, 20);
            this.numPingSheetRowNo.TabIndex = 37;
            this.numPingSheetRowNo.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label14.Location = new System.Drawing.Point(208, 25);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(67, 13);
            this.label14.TabIndex = 36;
            this.label14.Text = "Start Row:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkHideSeperator);
            this.groupBox1.Controls.Add(this.chkHideText);
            this.groupBox1.Controls.Add(this.chkLargeButtons);
            this.groupBox1.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(303, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(154, 78);
            this.groupBox1.TabIndex = 40;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Toolbar";
            // 
            // chkHideSeperator
            // 
            this.chkHideSeperator.AutoSize = true;
            this.chkHideSeperator.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideSeperator.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkHideSeperator.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHideSeperator.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideSeperator.Location = new System.Drawing.Point(22, 55);
            this.chkHideSeperator.Name = "chkHideSeperator";
            this.chkHideSeperator.Size = new System.Drawing.Size(119, 17);
            this.chkHideSeperator.TabIndex = 39;
            this.chkHideSeperator.Text = "Hide Seperators:";
            this.chkHideSeperator.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideSeperator.UseVisualStyleBackColor = true;
            // 
            // chkHideText
            // 
            this.chkHideText.AutoSize = true;
            this.chkHideText.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkHideText.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHideText.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkHideText.Location = new System.Drawing.Point(16, 36);
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
            this.chkLargeButtons.Location = new System.Drawing.Point(34, 19);
            this.chkLargeButtons.Name = "chkLargeButtons";
            this.chkLargeButtons.Size = new System.Drawing.Size(107, 17);
            this.chkLargeButtons.TabIndex = 37;
            this.chkLargeButtons.Text = "Large Buttons:";
            this.chkLargeButtons.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkLargeButtons.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkClearFormatting);
            this.groupBox2.Controls.Add(this.chkTurnOffScreenValidation);
            this.groupBox2.Controls.Add(this.chkDisplayTimeTaken);
            this.groupBox2.Controls.Add(this.chkProduceCompleteMessageBox);
            this.groupBox2.Controls.Add(this.chkProduceInitialMessageBox);
            this.groupBox2.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(5, 10);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(292, 126);
            this.groupBox2.TabIndex = 41;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Messages and Warnings";
            this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // chkClearFormatting
            // 
            this.chkClearFormatting.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkClearFormatting.Checked = true;
            this.chkClearFormatting.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkClearFormatting.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkClearFormatting.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkClearFormatting.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkClearFormatting.Location = new System.Drawing.Point(75, 100);
            this.chkClearFormatting.Name = "chkClearFormatting";
            this.chkClearFormatting.Size = new System.Drawing.Size(203, 17);
            this.chkClearFormatting.TabIndex = 49;
            this.chkClearFormatting.Text = "Selecting No Clears formatting";
            this.chkClearFormatting.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkClearFormatting.UseVisualStyleBackColor = true;
            // 
            // chkTurnOffScreenValidation
            // 
            this.chkTurnOffScreenValidation.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTurnOffScreenValidation.Checked = true;
            this.chkTurnOffScreenValidation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTurnOffScreenValidation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkTurnOffScreenValidation.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTurnOffScreenValidation.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTurnOffScreenValidation.Location = new System.Drawing.Point(99, 79);
            this.chkTurnOffScreenValidation.Name = "chkTurnOffScreenValidation";
            this.chkTurnOffScreenValidation.Size = new System.Drawing.Size(179, 17);
            this.chkTurnOffScreenValidation.TabIndex = 47;
            this.chkTurnOffScreenValidation.Text = "Turn off screen validation?";
            this.chkTurnOffScreenValidation.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTurnOffScreenValidation.UseVisualStyleBackColor = true;
            // 
            // chkDisplayTimeTaken
            // 
            this.chkDisplayTimeTaken.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Checked = true;
            this.chkDisplayTimeTaken.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDisplayTimeTaken.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkDisplayTimeTaken.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDisplayTimeTaken.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.Location = new System.Drawing.Point(141, 58);
            this.chkDisplayTimeTaken.Name = "chkDisplayTimeTaken";
            this.chkDisplayTimeTaken.Size = new System.Drawing.Size(137, 17);
            this.chkDisplayTimeTaken.TabIndex = 42;
            this.chkDisplayTimeTaken.Text = "Display Time Taken?";
            this.chkDisplayTimeTaken.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkDisplayTimeTaken.UseVisualStyleBackColor = true;
            // 
            // chkProduceCompleteMessageBox
            // 
            this.chkProduceCompleteMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceCompleteMessageBox.Checked = true;
            this.chkProduceCompleteMessageBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProduceCompleteMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceCompleteMessageBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceCompleteMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceCompleteMessageBox.Location = new System.Drawing.Point(51, 37);
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
            this.chkProduceInitialMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.Checked = true;
            this.chkProduceInitialMessageBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProduceInitialMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceInitialMessageBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceInitialMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.Location = new System.Drawing.Point(15, 16);
            this.chkProduceInitialMessageBox.Name = "chkProduceInitialMessageBox";
            this.chkProduceInitialMessageBox.Size = new System.Drawing.Size(263, 17);
            this.chkProduceInitialMessageBox.TabIndex = 40;
            this.chkProduceInitialMessageBox.Text = "Display Inital Confirmation Message Box?";
            this.chkProduceInitialMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceInitialMessageBox.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkNotFoundBold);
            this.groupBox3.Controls.Add(this.chkFoundBold);
            this.groupBox3.Controls.Add(this.chkLBColorOrCompare);
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
            this.groupBox3.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(5, 138);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(452, 90);
            this.groupBox3.TabIndex = 42;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Compare Sheets Options";
            // 
            // chkNotFoundBold
            // 
            this.chkNotFoundBold.AutoSize = true;
            this.chkNotFoundBold.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkNotFoundBold.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkNotFoundBold.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkNotFoundBold.Location = new System.Drawing.Point(385, 50);
            this.chkNotFoundBold.Name = "chkNotFoundBold";
            this.chkNotFoundBold.Size = new System.Drawing.Size(47, 17);
            this.chkNotFoundBold.TabIndex = 45;
            this.chkNotFoundBold.Text = "Bold";
            this.chkNotFoundBold.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkNotFoundBold.UseVisualStyleBackColor = true;
            this.chkNotFoundBold.CheckedChanged += new System.EventHandler(this.chkNotFoundBold_CheckedChanged);
            // 
            // chkFoundBold
            // 
            this.chkFoundBold.AutoSize = true;
            this.chkFoundBold.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkFoundBold.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkFoundBold.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkFoundBold.Location = new System.Drawing.Point(385, 23);
            this.chkFoundBold.Name = "chkFoundBold";
            this.chkFoundBold.Size = new System.Drawing.Size(47, 17);
            this.chkFoundBold.TabIndex = 44;
            this.chkFoundBold.Text = "Bold";
            this.chkFoundBold.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkFoundBold.UseVisualStyleBackColor = true;
            this.chkFoundBold.CheckedChanged += new System.EventHandler(this.chkFoundBold_CheckedChanged);
            // 
            // chkLBColorOrCompare
            // 
            this.chkLBColorOrCompare.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.chkLBColorOrCompare.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.chkLBColorOrCompare.CheckOnClick = true;
            this.chkLBColorOrCompare.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLBColorOrCompare.FormattingEnabled = true;
            this.chkLBColorOrCompare.Items.AddRange(new object[] {
            "Colour",
            "Clear"});
            this.chkLBColorOrCompare.Location = new System.Drawing.Point(99, 16);
            this.chkLBColorOrCompare.Name = "chkLBColorOrCompare";
            this.chkLBColorOrCompare.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkLBColorOrCompare.Size = new System.Drawing.Size(63, 32);
            this.chkLBColorOrCompare.TabIndex = 43;
            this.chkLBColorOrCompare.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkLBColorOrCompare_ItemCheck);
            this.chkLBColorOrCompare.SelectedIndexChanged += new System.EventHandler(this.chkLBColorOrCompare_SelectedIndexChanged_1);
            this.chkLBColorOrCompare.Leave += new System.EventHandler(this.chkLBColorOrCompare_Leave);
            // 
            // btnColourNotFoundBack
            // 
            this.btnColourNotFoundBack.Location = new System.Drawing.Point(338, 50);
            this.btnColourNotFoundBack.Name = "btnColourNotFoundBack";
            this.btnColourNotFoundBack.Size = new System.Drawing.Size(40, 19);
            this.btnColourNotFoundBack.TabIndex = 37;
            this.btnColourNotFoundBack.Text = "Back";
            this.btnColourNotFoundBack.UseVisualStyleBackColor = true;
            this.btnColourNotFoundBack.Click += new System.EventHandler(this.btnColourNotFoundBack_Click);
            // 
            // btnColourFoundBack
            // 
            this.btnColourFoundBack.Location = new System.Drawing.Point(339, 21);
            this.btnColourFoundBack.Name = "btnColourFoundBack";
            this.btnColourFoundBack.Size = new System.Drawing.Size(40, 19);
            this.btnColourFoundBack.TabIndex = 36;
            this.btnColourFoundBack.Text = "Back";
            this.btnColourFoundBack.UseVisualStyleBackColor = true;
            this.btnColourFoundBack.Click += new System.EventHandler(this.btnColourFoundBack_Click);
            // 
            // btnColourFound
            // 
            this.btnColourFound.Location = new System.Drawing.Point(293, 21);
            this.btnColourFound.Name = "btnColourFound";
            this.btnColourFound.Size = new System.Drawing.Size(40, 19);
            this.btnColourFound.TabIndex = 35;
            this.btnColourFound.Text = "Fore";
            this.btnColourFound.UseVisualStyleBackColor = true;
            this.btnColourFound.Click += new System.EventHandler(this.btnColourFound_Click);
            // 
            // btnColourNotFound
            // 
            this.btnColourNotFound.Location = new System.Drawing.Point(293, 50);
            this.btnColourNotFound.Name = "btnColourNotFound";
            this.btnColourNotFound.Size = new System.Drawing.Size(39, 19);
            this.btnColourNotFound.TabIndex = 34;
            this.btnColourNotFound.Text = "Fore";
            this.btnColourNotFound.UseVisualStyleBackColor = true;
            this.btnColourNotFound.Click += new System.EventHandler(this.btnColourNotFound_Click);
            // 
            // txtColourNotFound
            // 
            this.txtColourNotFound.BackColor = System.Drawing.SystemColors.Window;
            this.txtColourNotFound.Enabled = false;
            this.txtColourNotFound.ForeColor = System.Drawing.Color.Blue;
            this.txtColourNotFound.Location = new System.Drawing.Point(217, 49);
            this.txtColourNotFound.Name = "txtColourNotFound";
            this.txtColourNotFound.ReadOnly = true;
            this.txtColourNotFound.Size = new System.Drawing.Size(70, 20);
            this.txtColourNotFound.TabIndex = 33;
            this.txtColourNotFound.Text = "Not Found";
            this.txtColourNotFound.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtColourNotFound.Click += new System.EventHandler(this.btnColourNotFound_Click);
            // 
            // txtColourFound
            // 
            this.txtColourFound.BackColor = System.Drawing.SystemColors.Window;
            this.txtColourFound.Enabled = false;
            this.txtColourFound.ForeColor = System.Drawing.Color.Red;
            this.txtColourFound.Location = new System.Drawing.Point(217, 20);
            this.txtColourFound.Name = "txtColourFound";
            this.txtColourFound.ReadOnly = true;
            this.txtColourFound.Size = new System.Drawing.Size(70, 20);
            this.txtColourFound.TabIndex = 32;
            this.txtColourFound.Text = "Found";
            this.txtColourFound.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtColourFound.Click += new System.EventHandler(this.btnColourFound_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(168, 53);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(43, 13);
            this.label9.TabIndex = 31;
            this.label9.Text = "Items:";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(168, 24);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(43, 13);
            this.label8.TabIndex = 30;
            this.label8.Text = "Items:";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // numComparingStartRow
            // 
            this.numComparingStartRow.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numComparingStartRow.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numComparingStartRow.Location = new System.Drawing.Point(99, 56);
            this.numComparingStartRow.Name = "numComparingStartRow";
            this.numComparingStartRow.Size = new System.Drawing.Size(63, 20);
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
            this.label5.Location = new System.Drawing.Point(9, 56);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 13);
            this.label5.TabIndex = 28;
            this.label5.Text = "Starting Row:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 27;
            this.label1.Text = "Differences:";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.numDupliateColumnToCheck);
            this.groupBox4.Controls.Add(this.label6);
            this.groupBox4.Controls.Add(this.numNoOfColumnsToCheck);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.chkLBColourOrDelete);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(5, 231);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(452, 93);
            this.groupBox4.TabIndex = 45;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Duplicate Rows Options";
            // 
            // numDupliateColumnToCheck
            // 
            this.numDupliateColumnToCheck.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numDupliateColumnToCheck.Location = new System.Drawing.Point(320, 19);
            this.numDupliateColumnToCheck.Name = "numDupliateColumnToCheck";
            this.numDupliateColumnToCheck.Size = new System.Drawing.Size(52, 20);
            this.numDupliateColumnToCheck.TabIndex = 52;
            this.numDupliateColumnToCheck.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label6.Location = new System.Drawing.Point(229, 21);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 13);
            this.label6.TabIndex = 51;
            this.label6.Text = "Start Column:";
            // 
            // numNoOfColumnsToCheck
            // 
            this.numNoOfColumnsToCheck.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numNoOfColumnsToCheck.Location = new System.Drawing.Point(321, 49);
            this.numNoOfColumnsToCheck.Name = "numNoOfColumnsToCheck";
            this.numNoOfColumnsToCheck.Size = new System.Drawing.Size(52, 20);
            this.numNoOfColumnsToCheck.TabIndex = 50;
            this.numNoOfColumnsToCheck.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label4.Location = new System.Drawing.Point(169, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(145, 13);
            this.label4.TabIndex = 49;
            this.label4.Text = "No Of Columns To Check:";
            // 
            // chkLBColourOrDelete
            // 
            this.chkLBColourOrDelete.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.chkLBColourOrDelete.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.chkLBColourOrDelete.CheckOnClick = true;
            this.chkLBColourOrDelete.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLBColourOrDelete.FormattingEnabled = true;
            this.chkLBColourOrDelete.Items.AddRange(new object[] {
            "Colour",
            "Delete",
            "Clear"});
            this.chkLBColourOrDelete.Location = new System.Drawing.Point(75, 26);
            this.chkLBColourOrDelete.Name = "chkLBColourOrDelete";
            this.chkLBColourOrDelete.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkLBColourOrDelete.Size = new System.Drawing.Size(80, 47);
            this.chkLBColourOrDelete.TabIndex = 46;
            this.chkLBColourOrDelete.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkLBColourOrDelete_ItemCheck);
            this.chkLBColourOrDelete.Leave += new System.EventHandler(this.chkLBColourOrDelete_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(18, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 45;
            this.label3.Text = "Do what?";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.numTimeSheetRowNo);
            this.groupBox5.Controls.Add(this.label11);
            this.groupBox5.Controls.Add(this.chkTimeSheetGetRowNo);
            this.groupBox5.Location = new System.Drawing.Point(5, 328);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(452, 45);
            this.groupBox5.TabIndex = 47;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Time Sheet";
            // 
            // numTimeSheetRowNo
            // 
            this.numTimeSheetRowNo.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numTimeSheetRowNo.Location = new System.Drawing.Point(348, 19);
            this.numTimeSheetRowNo.Maximum = new decimal(new int[] {
            300000,
            0,
            0,
            0});
            this.numTimeSheetRowNo.Name = "numTimeSheetRowNo";
            this.numTimeSheetRowNo.Size = new System.Drawing.Size(52, 20);
            this.numTimeSheetRowNo.TabIndex = 31;
            this.numTimeSheetRowNo.Value = new decimal(new int[] {
            3754,
            0,
            0,
            0});
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label11.Location = new System.Drawing.Point(196, 21);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(145, 13);
            this.label11.TabIndex = 30;
            this.label11.Text = "Timesheet Row Start No:";
            // 
            // chkTimeSheetGetRowNo
            // 
            this.chkTimeSheetGetRowNo.AutoSize = true;
            this.chkTimeSheetGetRowNo.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkTimeSheetGetRowNo.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.chkTimeSheetGetRowNo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.Location = new System.Drawing.Point(16, 19);
            this.chkTimeSheetGetRowNo.Name = "chkTimeSheetGetRowNo";
            this.chkTimeSheetGetRowNo.Size = new System.Drawing.Size(119, 17);
            this.chkTimeSheetGetRowNo.TabIndex = 29;
            this.chkTimeSheetGetRowNo.Text = "Get next row No:";
            this.chkTimeSheetGetRowNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTimeSheetGetRowNo.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.numZapSheetStartRow);
            this.groupBox7.Controls.Add(this.label16);
            this.groupBox7.Controls.Add(this.numColNoForExtractedName);
            this.groupBox7.Controls.Add(this.label15);
            this.groupBox7.Controls.Add(this.chkExtractFileName);
            this.groupBox7.Controls.Add(this.label2);
            this.groupBox7.Controls.Add(this.cmboWhichDate);
            this.groupBox7.Controls.Add(this.chkTestCode);
            this.groupBox7.Controls.Add(this.label7);
            this.groupBox7.Controls.Add(this.cmboDelMode);
            this.groupBox7.Controls.Add(this.label10);
            this.groupBox7.Controls.Add(this.numHighlightRowsOver);
            this.groupBox7.Location = new System.Drawing.Point(5, 459);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(452, 178);
            this.groupBox7.TabIndex = 50;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Misc";
            // 
            // numZapSheetStartRow
            // 
            this.numZapSheetStartRow.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numZapSheetStartRow.Location = new System.Drawing.Point(178, 150);
            this.numZapSheetStartRow.Name = "numZapSheetStartRow";
            this.numZapSheetStartRow.Size = new System.Drawing.Size(52, 20);
            this.numZapSheetStartRow.TabIndex = 57;
            this.numZapSheetStartRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label16.Location = new System.Drawing.Point(45, 152);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(127, 13);
            this.label16.TabIndex = 56;
            this.label16.Text = "Zap Sheet Start Row:";
            // 
            // numColNoForExtractedName
            // 
            this.numColNoForExtractedName.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numColNoForExtractedName.Location = new System.Drawing.Point(178, 120);
            this.numColNoForExtractedName.Name = "numColNoForExtractedName";
            this.numColNoForExtractedName.Size = new System.Drawing.Size(52, 20);
            this.numColNoForExtractedName.TabIndex = 55;
            this.numColNoForExtractedName.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.label15.Location = new System.Drawing.Point(33, 122);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(139, 13);
            this.label15.TabIndex = 54;
            this.label15.Text = "Extracted Name Column:";
            // 
            // chkExtractFileName
            // 
            this.chkExtractFileName.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkExtractFileName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkExtractFileName.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkExtractFileName.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkExtractFileName.Location = new System.Drawing.Point(12, 97);
            this.chkExtractFileName.Name = "chkExtractFileName";
            this.chkExtractFileName.Size = new System.Drawing.Size(186, 17);
            this.chkExtractFileName.TabIndex = 53;
            this.chkExtractFileName.Text = "Extract FileName from Path?";
            this.chkExtractFileName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkExtractFileName.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 13);
            this.label2.TabIndex = 52;
            this.label2.Text = "Which Date Time from file";
            // 
            // cmboWhichDate
            // 
            this.cmboWhichDate.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cmboWhichDate.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboWhichDate.FormattingEnabled = true;
            this.cmboWhichDate.Items.AddRange(new object[] {
            "CreationTime",
            "CreationTimeUtc",
            "LastAccessTime",
            "LastAccessTimeUtc",
            "LastWriteTime",
            "LastWriteTimeUtc"});
            this.cmboWhichDate.Location = new System.Drawing.Point(178, 68);
            this.cmboWhichDate.Name = "cmboWhichDate";
            this.cmboWhichDate.Size = new System.Drawing.Size(139, 23);
            this.cmboWhichDate.TabIndex = 51;
            // 
            // chkTestCode
            // 
            this.chkTestCode.AutoSize = true;
            this.chkTestCode.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTestCode.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkTestCode.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkTestCode.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTestCode.Location = new System.Drawing.Point(240, 45);
            this.chkTestCode.Name = "chkTestCode";
            this.chkTestCode.Size = new System.Drawing.Size(77, 17);
            this.chkTestCode.TabIndex = 40;
            this.chkTestCode.Text = "Test Code";
            this.chkTestCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTestCode.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(9, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(163, 13);
            this.label7.TabIndex = 36;
            this.label7.Text = "Del blank Lines which Mode";
            // 
            // cmboDelMode
            // 
            this.cmboDelMode.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.cmboDelMode.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDelMode.FormattingEnabled = true;
            this.cmboDelMode.Items.AddRange(new object[] {
            "Mode: A",
            "Mode: B",
            "Mode: C",
            "Mode: D"});
            this.cmboDelMode.Location = new System.Drawing.Point(178, 13);
            this.cmboDelMode.Name = "cmboDelMode";
            this.cmboDelMode.Size = new System.Drawing.Size(139, 23);
            this.cmboDelMode.TabIndex = 35;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(27, 44);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(145, 13);
            this.label10.TabIndex = 37;
            this.label10.Text = "Highlight lengths over:";
            // 
            // numHighlightRowsOver
            // 
            this.numHighlightRowsOver.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.numHighlightRowsOver.Location = new System.Drawing.Point(178, 42);
            this.numHighlightRowsOver.Maximum = new decimal(new int[] {
            1024,
            0,
            0,
            0});
            this.numHighlightRowsOver.Name = "numHighlightRowsOver";
            this.numHighlightRowsOver.Size = new System.Drawing.Size(52, 20);
            this.numHighlightRowsOver.TabIndex = 38;
            this.numHighlightRowsOver.Value = new decimal(new int[] {
            180,
            0,
            0,
            0});
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.btnTestCode);
            this.groupBox8.Controls.Add(this.btnApply);
            this.groupBox8.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox8.Location = new System.Drawing.Point(0, 678);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(460, 50);
            this.groupBox8.TabIndex = 51;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Foot";
            // 
            // btnApply
            // 
            this.btnApply.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnApply.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnApply.Location = new System.Drawing.Point(348, 19);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(100, 23);
            this.btnApply.TabIndex = 39;
            this.btnApply.Text = "&Apply / Close";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnTestCode
            // 
            this.btnTestCode.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnTestCode.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnTestCode.Location = new System.Drawing.Point(204, 19);
            this.btnTestCode.Name = "btnTestCode";
            this.btnTestCode.Size = new System.Drawing.Size(100, 23);
            this.btnTestCode.TabIndex = 40;
            this.btnTestCode.Text = "&Test Code";
            this.btnTestCode.UseVisualStyleBackColor = true;
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(460, 728);
            this.Controls.Add(this.groupBox8);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmSettings";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings ...";
            this.Load += new System.EventHandler(this.Form1_Load);
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingWrite)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColPingRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numPingSheetRowNo)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numComparingStartRow)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDupliateColumnToCheck)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numNoOfColumnsToCheck)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTimeSheetRowNo)).EndInit();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numZapSheetStartRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColNoForExtractedName)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numHighlightRowsOver)).EndInit();
            this.groupBox8.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ColorDialog colorDialog1;
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
        private System.Windows.Forms.Button btnColourNotFound;
        private System.Windows.Forms.Button btnColourFound;
        private System.Windows.Forms.Button btnColourNotFoundBack;
        private System.Windows.Forms.Button btnColourFoundBack;
        private System.Windows.Forms.CheckedListBox chkLBColorOrCompare;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.NumericUpDown numDupliateColumnToCheck;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown numNoOfColumnsToCheck;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckedListBox chkLBColourOrDelete;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.CheckBox chkHideSeperator;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.NumericUpDown numTimeSheetRowNo;
        private System.Windows.Forms.Label label11;
        public System.Windows.Forms.CheckBox chkTimeSheetGetRowNo;
        public System.Windows.Forms.CheckBox chkClearFormatting;
        public System.Windows.Forms.CheckBox chkTurnOffScreenValidation;
        private System.Windows.Forms.NumericUpDown numColPingWrite;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.NumericUpDown numColPingRead;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.NumericUpDown numPingSheetRowNo;
        private System.Windows.Forms.Label label14;
        public System.Windows.Forms.CheckBox chkNotFoundBold;
        public System.Windows.Forms.CheckBox chkFoundBold;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmboDelMode;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown numHighlightRowsOver;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Button btnApply;
        public System.Windows.Forms.CheckBox chkTestCode;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmboWhichDate;
        public System.Windows.Forms.CheckBox chkExtractFileName;
        private System.Windows.Forms.NumericUpDown numColNoForExtractedName;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.NumericUpDown numZapSheetStartRow;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Button btnTestCode;
    }
}