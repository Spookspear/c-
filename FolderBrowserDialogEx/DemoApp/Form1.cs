using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DaveChambers.FolderBrowserDialogEx;

namespace DemoApp
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void btnStd_Click(object sender, EventArgs e)
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			fbd.SelectedPath = txtInitialFolder.Text;
			fbd.ShowNewFolderButton = chkShowNewFolderBtn.Checked;
			fbd.RootFolder = (Environment.SpecialFolder)comboRootFolder.SelectedValue;

			DialogResult dr = fbd.ShowDialog();
			txtRtnCode.Text = dr.ToString();
			txtRtnFolder.Text = "";
			if(dr == DialogResult.OK)
			{
				txtRtnFolder.Text = fbd.SelectedPath;
			}
		}

		private void btnCustom_Click(object sender, EventArgs e)
		{
			FolderBrowserDialogEx cfbd = new FolderBrowserDialogEx();
			cfbd.Title = txtCaption.Text;
			cfbd.SelectedPath = txtInitialFolder.Text;
			cfbd.ShowEditbox = chkEditbox.Checked;
			cfbd.ShowNewFolderButton = chkShowNewFolderBtn.Checked;
			cfbd.RootFolder = (Environment.SpecialFolder)comboRootFolder.SelectedValue;
			cfbd.StartPosition = FormStartPosition.CenterScreen;

			DialogResult dr = cfbd.ShowDialog(this);
			txtRtnCode.Text = dr.ToString();
			txtRtnFolder.Text = "";
			if (dr == DialogResult.OK)
			{
				txtRtnFolder.Text = cfbd.SelectedPath;
			}
		}

		private void Form1_Load(object sender, EventArgs e)
		{
			//_fillSpecialFolderCombo(comboInitialFolder);
			//_fillSpecialFolderCombo(comboRootFolder);
			comboStartPos.DataSource = Enum.GetValues(typeof(FormStartPosition));
			comboRootFolder.DataSource = Enum.GetValues(typeof(Environment.SpecialFolder));
		}

		//private void _fillSpecialFolderCombo(ComboBox combo)
		//{
		//    //Environment.SpecialFolder.CommonApplicationData
		//    combo.DataSource = Enum.GetValues(typeof(Environment.SpecialFolder));
		//}
	}
}
