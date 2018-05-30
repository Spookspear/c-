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
            this.btnCancel = new System.Windows.Forms.Button();
            this.chkProduceMessageBox = new System.Windows.Forms.CheckBox();
            this.cmboDifferences = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCancel.Location = new System.Drawing.Point(390, 279);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Exit";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // chkProduceMessageBox
            // 
            this.chkProduceMessageBox.AutoSize = true;
            this.chkProduceMessageBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.chkProduceMessageBox.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkProduceMessageBox.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.Location = new System.Drawing.Point(12, 12);
            this.chkProduceMessageBox.Name = "chkProduceMessageBox";
            this.chkProduceMessageBox.Size = new System.Drawing.Size(163, 19);
            this.chkProduceMessageBox.TabIndex = 1;
            this.chkProduceMessageBox.Text = "Produce Message Box?";
            this.chkProduceMessageBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkProduceMessageBox.UseVisualStyleBackColor = true;
            // 
            // cmboDifferences
            // 
            this.cmboDifferences.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmboDifferences.FormattingEnabled = true;
            this.cmboDifferences.Location = new System.Drawing.Point(166, 37);
            this.cmboDifferences.Name = "cmboDifferences";
            this.cmboDifferences.Size = new System.Drawing.Size(139, 23);
            this.cmboDifferences.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(76, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "Differences";
            // 
            // frmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(477, 314);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmboDifferences);
            this.Controls.Add(this.chkProduceMessageBox);
            this.Controls.Add(this.btnCancel);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.MinimizeBox = false;
            this.Name = "frmSettings";
            this.ShowInTaskbar = false;
            this.Text = "Settings ...";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        public System.Windows.Forms.CheckBox chkProduceMessageBox;
        private System.Windows.Forms.ComboBox cmboDifferences;
        private System.Windows.Forms.Label label1;
    }
}