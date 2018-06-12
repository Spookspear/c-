#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Tools.Ribbon;

using ToolbarOfFunctions_CommonClasses;
using ToolbarOfFunctions;

using System.Xml.Serialization;
using System.IO;


namespace ToolbarOfFunctions
{
    public partial class frmSettings : Form
    {

        InformationForSettingsForm myData = new InformationForSettingsForm();

        public frmSettings()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            // InformationForSettingsForm myData = new InformationForSettingsForm();
            myData = SaveXML.LoadData();

            chkLargeButtons.Checked = myData.LargeButtons;
            chkHideText.Checked = myData.HideText;
            // cmboCompareDifferences.Text = myData.CompareOrColour;

            chkLBColorOrCompare.Text = myData.CompareOrColourNew;

            // tick the correct box
            if (chkLBColorOrCompare.CheckedItems.Count == 0)
            {

                if (myData.CompareOrColourNew == "Colour")
                {
                    chkLBColorOrCompare.SetItemChecked(0, true);
                }
                else
                {
                    chkLBColorOrCompare.SetItemChecked(1, true);
                }
            }

            // - New 
            chkLBHighLightOrDelete.Text = myData.HighLightOrDeleteNew;

            // tick the correct box
            if (chkLBHighLightOrDelete.CheckedItems.Count == 0)
            {

                if (myData.HighLightOrDeleteNew == "Highlight")
                {
                    chkLBHighLightOrDelete.SetItemChecked(0, true);
                }
                else if (myData.HighLightOrDeleteNew == "Delete")
                {
                    chkLBHighLightOrDelete.SetItemChecked(1, true);
                }
                else if (myData.HighLightOrDeleteNew == "Clear")
                {
                    chkLBHighLightOrDelete.SetItemChecked(2, true);
                }

            }

            // -New eof


            // cmboHighLightOrDelete.Text = myData.HighLightOrDelete;
            chkDisplayTimeTaken.Checked = myData.DisplayTimeTaken;

            chkProduceInitialMessageBox.Checked = myData.ProduceInitialMessageBox;
            chkProduceCompleteMessageBox.Checked = myData.ProduceCompleteMessageBox;

            cmboDelModeAorBorC.Text = myData.DelModeAorBorC;
            numHighlightRowsOver.Value = myData.HighlightRowsOver;
            // ---------

            numNoOfColumnsToCheck.Value = myData.NoOfColumnsToCheck;
            numComparingStartRow.Value = myData.ComparingStartRow;
            numDupliateColumnToCheck.Value = myData.DupliateColumnToCheck;

            // txtColourFound.Text = myData.ColourFoundText;
            // txtColourNotFound.Text = myData.ColourNotFoundText;

            txtColourFound.ForeColor = ColorTranslator.FromHtml(myData.ColourFore_Found);
            txtColourNotFound.ForeColor = ColorTranslator.FromHtml(myData.ColourFore_NotFound);

            txtColourFound.BackColor = ColorTranslator.FromHtml(myData.ColourBack_Found);
            txtColourNotFound.BackColor = ColorTranslator.FromHtml(myData.ColourBack_NotFound);


            // ---------
            numTimeSheetRowNo.Value = myData.TimeSheetRowNo;
            chkTimeSheetGetRowNo.Checked = myData.TimeSheetGetRowNo;

            // ---------
            numPingSheetRowNo.Value = myData.PingSheetRowNo;
            numColPingRead.Value = myData.ColPingRead;
            numColPingWrite.Value = myData.ColPingWrite;

            chkTestCode.Checked = myData.TestCode;

            chkRecordTimes.Checked = myData.RecordTimes;

            /*
            if (cmboCompareDifferences.Items.Count != 2)
            {
                cmboCompareDifferences.Items.Add("Colour");
                cmboCompareDifferences.Items.Add("Clear");
                // cmboCompareDifferences.SelectedIndex = 0;
            }
            


            if (cmboHighLightOrDelete.Items.Count != 2)
            {
                cmboHighLightOrDelete.Items.Add("Highlight");
                cmboHighLightOrDelete.Items.Add("Delete");
                // cmboHighLightOrDelete.SelectedIndex = 0;
            }
            */

            if (cmboDelModeAorBorC.Items.Count != 3)
            {
                cmboDelModeAorBorC.Items.Add("Mode: A");
                cmboDelModeAorBorC.Items.Add("Mode: B");
                cmboDelModeAorBorC.Items.Add("Mode: C");
                // cmboDelModeAorBorC.SelectedIndex = 0;

            }

            checkCompareCombo();

        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            myData.LargeButtons = chkLargeButtons.Checked;
            myData.HideText = chkHideText.Checked;

            // myData.CompareOrColour = cmboCompareDifferences.Text;
            myData.CompareOrColourNew = chkLBColorOrCompare.Text;               // new item

            // myData.HighLightOrDelete = cmboHighLightOrDelete.Text;
            myData.HighLightOrDeleteNew = chkLBHighLightOrDelete.Text;          // new item

            myData.DisplayTimeTaken = chkDisplayTimeTaken.Checked;

            myData.ProduceInitialMessageBox = chkProduceInitialMessageBox.Checked;
            myData.ProduceCompleteMessageBox = chkProduceCompleteMessageBox.Checked;

            myData.DelModeAorBorC = cmboDelModeAorBorC.Text;
            myData.HighlightRowsOver = numHighlightRowsOver.Value;
            // ---------

            myData.NoOfColumnsToCheck = numNoOfColumnsToCheck.Value;
            myData.ComparingStartRow = numComparingStartRow.Value;
            myData.DupliateColumnToCheck = numDupliateColumnToCheck.Value;

            // myData.ColourFoundText = txtColourFound.Text;
            // myData.ColourNotFoundText = txtColourNotFound.Text;

            myData.ColourFore_Found = ColorTranslator.ToHtml(txtColourFound.ForeColor);
            myData.ColourFore_NotFound = ColorTranslator.ToHtml(txtColourNotFound.ForeColor);

            myData.ColourBack_Found = ColorTranslator.ToHtml(txtColourFound.BackColor);
            myData.ColourBack_NotFound = ColorTranslator.ToHtml(txtColourNotFound.BackColor);

            // ---------
            myData.TimeSheetRowNo = numTimeSheetRowNo.Value;
            myData.TimeSheetGetRowNo = chkTimeSheetGetRowNo.Checked;

            // ---------
            myData.PingSheetRowNo = numPingSheetRowNo.Value;
            myData.ColPingRead = numColPingRead.Value;
            myData.ColPingWrite = numColPingWrite.Value;

            //---- misc
            myData.TestCode = chkTestCode.Checked;
            myData.RecordTimes = chkRecordTimes.Checked;

            SaveXML.SaveData(myData);

            this.Hide();

        }

        private void btnColourNotFound_Click(object sender, EventArgs e)
        {

            // myData = SaveXML.LoadData();
            // colorDialog1.Color = ColorTranslator.FromHtml(myData.ColourFore_NotFound);

            colorDialog1.Color = txtColourNotFound.ForeColor;

            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                txtColourNotFound.ForeColor = colorDialog1.Color;

                // if (colorDialog1.Color.IsNamedColor)                {                }
                // txtColourNotFound.Text = colorDialog1.Color.Name;
            }

        }

        private void btnColourFound_Click(object sender, EventArgs e)
        {
            // myData = SaveXML.LoadData();
            // colorDialog1.Color = ColorTranslator.FromHtml(myData.ColourFore_Found);

            colorDialog1.Color = txtColourFound.ForeColor;

            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                txtColourFound.ForeColor = colorDialog1.Color;
                // txtColourFound.Text = colorDialog1.Color.Name;
                // if (colorDialog1.Color.IsNamedColor)                {                }
            }


        }


        private void btnColourFoundBack_Click(object sender, EventArgs e)
        {

            // myData = SaveXML.LoadData();
            // colorDialog1.Color = ColorTranslator.FromHtml(myData.ColourBack_Found);

            colorDialog1.Color = txtColourFound.BackColor;

            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                txtColourFound.BackColor = colorDialog1.Color;
            }

        }

        private void btnColourNotFoundBack_Click(object sender, EventArgs e)
        {
            // myData = SaveXML.LoadData();
            // colorDialog1.Color = ColorTranslator.FromHtml(myData.ColourBack_NotFound);

            colorDialog1.Color = txtColourNotFound.BackColor;

            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                txtColourNotFound.BackColor = colorDialog1.Color;
            }

        }


        private void cmboCompareDifferences_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkCompareCombo();
        }


        private void checkCompareCombo()
        {
            bool boolEnabled = false;

            if (chkLBColorOrCompare.Text == "Colour")
            {
                boolEnabled = true;
            }

            label8.Enabled = boolEnabled;
            txtColourFound.Enabled = boolEnabled;
            btnColourFound.Enabled = boolEnabled;
            label9.Enabled = boolEnabled;
            txtColourNotFound.Enabled = boolEnabled;
            btnColourNotFound.Enabled = boolEnabled;

            btnColourFoundBack.Enabled = boolEnabled;
            btnColourNotFoundBack.Enabled = boolEnabled;

        }

        private void chkProduceCompleteMessageBox_CheckedChanged(object sender, EventArgs e)
        {
            chkDisplayTimeTaken.Enabled = chkProduceCompleteMessageBox.Checked;
            chkDisplayTimeTaken.Checked = chkProduceCompleteMessageBox.Checked;
        }



        private void chkLBColorOrCompare_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked && chkLBColorOrCompare.CheckedItems.Count > 0)
            {
                chkLBColorOrCompare.ItemCheck -= chkLBColorOrCompare_ItemCheck;
                chkLBColorOrCompare.SetItemChecked(chkLBColorOrCompare.CheckedIndices[0], false);
                chkLBColorOrCompare.ItemCheck += chkLBColorOrCompare_ItemCheck;
            }

        }


        private void chkLBColorOrCompare_Leave(object sender, EventArgs e)
        {
            if (chkLBColorOrCompare.CheckedItems.Count == 0) 
                chkLBColorOrCompare.SetItemChecked(0, true);
            
        }


        private void chkLBColorOrCompare_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            checkCompareCombo();
        }

        // ----------

        private void chkLBHighLightOrDelete_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked && chkLBHighLightOrDelete.CheckedItems.Count > 0)
            {
                chkLBHighLightOrDelete.ItemCheck -= chkLBHighLightOrDelete_ItemCheck;
                chkLBHighLightOrDelete.SetItemChecked(chkLBHighLightOrDelete.CheckedIndices[0], false);
                chkLBHighLightOrDelete.ItemCheck += chkLBHighLightOrDelete_ItemCheck;
            }

        }


        private void chkLBHighLightOrDelete_Leave(object sender, EventArgs e)
        {
            if (chkLBHighLightOrDelete.CheckedItems.Count == 0)
                chkLBHighLightOrDelete.SetItemChecked(0, true);

        }


        // ----------




    }

}
