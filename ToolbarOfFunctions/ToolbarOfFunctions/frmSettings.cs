﻿#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolbarOfFunctions
{
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmboDifferences.Items.Add("Colour");
            cmboDifferences.Items.Add("Clear");
            cmboDifferences.SelectedIndex = 0;


        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

            // can I dynamically change the text if a button?
            
            this.Hide();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cmboDifferences_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


    }
}
