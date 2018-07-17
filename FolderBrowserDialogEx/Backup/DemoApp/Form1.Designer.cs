namespace DemoApp
{
	partial class Form1
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
			this.btnStd = new System.Windows.Forms.Button();
			this.btnCustom = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkEditbox = new System.Windows.Forms.CheckBox();
			this.label5 = new System.Windows.Forms.Label();
			this.comboStartPos = new System.Windows.Forms.ComboBox();
			this.txtCaption = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtInitialFolder = new System.Windows.Forms.TextBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.chkShowNewFolderBtn = new System.Windows.Forms.CheckBox();
			this.label3 = new System.Windows.Forms.Label();
			this.comboRootFolder = new System.Windows.Forms.ComboBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.txtRtnFolder = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtRtnCode = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnStd
			// 
			this.btnStd.Location = new System.Drawing.Point(83, 29);
			this.btnStd.Name = "btnStd";
			this.btnStd.Size = new System.Drawing.Size(146, 40);
			this.btnStd.TabIndex = 0;
			this.btnStd.Text = "Standard FolderBrowserDialog";
			this.btnStd.UseVisualStyleBackColor = true;
			this.btnStd.Click += new System.EventHandler(this.btnStd_Click);
			// 
			// btnCustom
			// 
			this.btnCustom.Location = new System.Drawing.Point(305, 29);
			this.btnCustom.Name = "btnCustom";
			this.btnCustom.Size = new System.Drawing.Size(146, 40);
			this.btnCustom.TabIndex = 1;
			this.btnCustom.Text = "FolderBrowserEx";
			this.btnCustom.UseVisualStyleBackColor = true;
			this.btnCustom.Click += new System.EventHandler(this.btnCustom_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkEditbox);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.comboStartPos);
			this.groupBox1.Controls.Add(this.txtCaption);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(527, 108);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Options for FolderBrowserEx ONLY";
			// 
			// chkEditbox
			// 
			this.chkEditbox.AutoSize = true;
			this.chkEditbox.Location = new System.Drawing.Point(18, 79);
			this.chkEditbox.Name = "chkEditbox";
			this.chkEditbox.Size = new System.Drawing.Size(99, 17);
			this.chkEditbox.TabIndex = 7;
			this.chkEditbox.Text = "Include Editbox";
			this.chkEditbox.UseVisualStyleBackColor = true;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(30, 50);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(69, 13);
			this.label5.TabIndex = 5;
			this.label5.Text = "Start Position";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// comboStartPos
			// 
			this.comboStartPos.FormattingEnabled = true;
			this.comboStartPos.Location = new System.Drawing.Point(105, 47);
			this.comboStartPos.Name = "comboStartPos";
			this.comboStartPos.Size = new System.Drawing.Size(295, 21);
			this.comboStartPos.TabIndex = 6;
			// 
			// txtCaption
			// 
			this.txtCaption.Location = new System.Drawing.Point(105, 21);
			this.txtCaption.Name = "txtCaption";
			this.txtCaption.Size = new System.Drawing.Size(409, 20);
			this.txtCaption.TabIndex = 1;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(56, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 13);
			this.label2.TabIndex = 0;
			this.label2.Text = "Caption";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtInitialFolder
			// 
			this.txtInitialFolder.Location = new System.Drawing.Point(105, 23);
			this.txtInitialFolder.Name = "txtInitialFolder";
			this.txtInitialFolder.Size = new System.Drawing.Size(409, 20);
			this.txtInitialFolder.TabIndex = 1;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.chkShowNewFolderBtn);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.comboRootFolder);
			this.groupBox2.Controls.Add(this.txtInitialFolder);
			this.groupBox2.Location = new System.Drawing.Point(12, 126);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(527, 106);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Options for BOTH FolderBrowserEx and FolderBrowserDialog";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(36, 26);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(63, 13);
			this.label6.TabIndex = 0;
			this.label6.Text = "Initial Folder";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// chkShowNewFolderBtn
			// 
			this.chkShowNewFolderBtn.AutoSize = true;
			this.chkShowNewFolderBtn.Location = new System.Drawing.Point(18, 76);
			this.chkShowNewFolderBtn.Name = "chkShowNewFolderBtn";
			this.chkShowNewFolderBtn.Size = new System.Drawing.Size(154, 17);
			this.chkShowNewFolderBtn.TabIndex = 4;
			this.chkShowNewFolderBtn.Text = "Show \"New Folder\" Button";
			this.chkShowNewFolderBtn.UseVisualStyleBackColor = true;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(37, 52);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(62, 13);
			this.label3.TabIndex = 2;
			this.label3.Text = "Root Folder";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// comboRootFolder
			// 
			this.comboRootFolder.FormattingEnabled = true;
			this.comboRootFolder.Location = new System.Drawing.Point(105, 49);
			this.comboRootFolder.Name = "comboRootFolder";
			this.comboRootFolder.Size = new System.Drawing.Size(295, 21);
			this.comboRootFolder.TabIndex = 3;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.btnStd);
			this.groupBox3.Controls.Add(this.btnCustom);
			this.groupBox3.Location = new System.Drawing.Point(12, 238);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(527, 86);
			this.groupBox3.TabIndex = 2;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Dialog Type";
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.txtRtnFolder);
			this.groupBox4.Controls.Add(this.label4);
			this.groupBox4.Controls.Add(this.txtRtnCode);
			this.groupBox4.Controls.Add(this.label1);
			this.groupBox4.Location = new System.Drawing.Point(12, 330);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(527, 77);
			this.groupBox4.TabIndex = 3;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Results";
			// 
			// txtRtnFolder
			// 
			this.txtRtnFolder.Location = new System.Drawing.Point(105, 45);
			this.txtRtnFolder.Name = "txtRtnFolder";
			this.txtRtnFolder.ReadOnly = true;
			this.txtRtnFolder.Size = new System.Drawing.Size(409, 20);
			this.txtRtnFolder.TabIndex = 3;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(18, 48);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(81, 13);
			this.label4.TabIndex = 2;
			this.label4.Text = "Selected Folder";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtRtnCode
			// 
			this.txtRtnCode.Location = new System.Drawing.Point(105, 19);
			this.txtRtnCode.Name = "txtRtnCode";
			this.txtRtnCode.ReadOnly = true;
			this.txtRtnCode.Size = new System.Drawing.Size(409, 20);
			this.txtRtnCode.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 22);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(67, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Return Code";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(552, 417);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "CustomFolderBrowserDialog Demo";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox2.PerformLayout();
			this.groupBox3.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox4.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnStd;
		private System.Windows.Forms.Button btnCustom;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TextBox txtInitialFolder;
		private System.Windows.Forms.TextBox txtCaption;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.CheckBox chkEditbox;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ComboBox comboRootFolder;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.TextBox txtRtnFolder;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox txtRtnCode;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ComboBox comboStartPos;
		private System.Windows.Forms.CheckBox chkShowNewFolderBtn;
		private System.Windows.Forms.Label label6;
	}
}

