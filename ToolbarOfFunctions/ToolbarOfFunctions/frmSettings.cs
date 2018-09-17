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
            myData = myData.LoadMyData();               // read data from settings file

            chkLargeButtons.Checked = myData.LargeButtons;
            chkHideText.Checked = myData.HideText;

            chkLBColorOrCompare.Text = myData.CompareOrColour;

            // tick the correct box
            if (chkLBColorOrCompare.CheckedItems.Count == 0)
            {

                if (myData.CompareOrColour == "Colour")
                {
                    chkLBColorOrCompare.SetItemChecked(0, true);
                } else {
                    chkLBColorOrCompare.SetItemChecked(1, true);
                }
            }

            // - New 
            chkLBColourOrDelete.Text = myData.ColourOrDelete;

            // tick the correct box
            if (chkLBColourOrDelete.CheckedItems.Count == 0)
            {
                if (myData.ColourOrDelete == "Colour")
                {
                    chkLBColourOrDelete.SetItemChecked(0, true);
                }
                else if (myData.ColourOrDelete == "Delete")
                {
                    chkLBColourOrDelete.SetItemChecked(1, true);
                }
                else if (myData.ColourOrDelete == "Clear")
                {
                    chkLBColourOrDelete.SetItemChecked(2, true);
                }

            }

            // -New eof

            chkDisplayTimeTaken.Checked = myData.DisplayTimeTaken;

            chkProduceInitialMessageBox.Checked = myData.ProduceInitialMessageBox;
            chkProduceCompleteMessageBox.Checked = myData.ProduceCompleteMessageBox;

            cmboDelMode.Text = myData.DelModeAorBorC;
            numHighlightRowsOver.Value = myData.HighlightRowsOver;
            // ---------

            numNoOfColumnsToCheck.Value = myData.NoOfColumnsToCheck;
            numComparingStartRow.Value = myData.ComparingStartRow;
            numDupliateColumnToCheck.Value = myData.DupliateColumnToCheck;

            txtColourFound.ForeColor = ColorTranslator.FromHtml(myData.ColourFore_Found);
            txtColourNotFound.ForeColor = ColorTranslator.FromHtml(myData.ColourFore_NotFound);

            chkFoundBold.Checked = myData.ColourBold_Found;

            txtColourFound.BackColor = ColorTranslator.FromHtml(myData.ColourBack_Found);
            txtColourNotFound.BackColor = ColorTranslator.FromHtml(myData.ColourBack_NotFound);

            chkNotFoundBold.Checked = myData.ColourBold_NotFound;

            // ---------
            numTimeSheetRowNo.Value = myData.TimeSheetRowNo;
            chkTimeSheetGetRowNo.Checked = myData.TimeSheetGetRowNo;

            // ---------
            numPingSheetRowNo.Value = myData.PingSheetRowNo;
            numColPingRead.Value = myData.ColPingRead;
            numColPingWrite.Value = myData.ColPingWrite;

            chkTestCode.Checked = myData.TestCode;

            chkTurnOffScreenValidation.Checked = myData.TurnOffScreenValidation;

            if (cmboDelMode.Items.Count != 4)
            {
                cmboDelMode.Items.Add("Mode: A");
                cmboDelMode.Items.Add("Mode: B");
                cmboDelMode.Items.Add("Mode: C");
                cmboDelMode.Items.Add("Mode: D");
            }

            checkCompareCombo();

            chkClearFormatting.Checked = myData.ClearFormatting;

            cmboWhichDate.Text = myData.FileDateTime;

            chkExtractFileName.Checked = myData.ExtractFileName;

            numColNoForExtractedName.Value = myData.ColExtractedFile;

            numZapSheetStartRow.Value = myData.ZapStartDefaultRow;

        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            myData.LargeButtons = chkLargeButtons.Checked;
            myData.HideText = chkHideText.Checked;

            myData.CompareOrColour = chkLBColorOrCompare.Text;               // new item
            myData.ColourOrDelete = chkLBColourOrDelete.Text;          // new item

            myData.DisplayTimeTaken = chkDisplayTimeTaken.Checked;

            myData.ProduceInitialMessageBox = chkProduceInitialMessageBox.Checked;
            myData.ProduceCompleteMessageBox = chkProduceCompleteMessageBox.Checked;

            myData.DelModeAorBorC = cmboDelMode.Text;
            myData.HighlightRowsOver = numHighlightRowsOver.Value;
            // ---------

            myData.NoOfColumnsToCheck = numNoOfColumnsToCheck.Value;
            myData.ComparingStartRow = numComparingStartRow.Value;
            myData.DupliateColumnToCheck = numDupliateColumnToCheck.Value;

            myData.ColourFore_Found = ColorTranslator.ToHtml(txtColourFound.ForeColor);
            myData.ColourFore_NotFound = ColorTranslator.ToHtml(txtColourNotFound.ForeColor);

            myData.ColourBold_Found = chkFoundBold.Checked;

            myData.ColourBack_Found = ColorTranslator.ToHtml(txtColourFound.BackColor);
            myData.ColourBack_NotFound = ColorTranslator.ToHtml(txtColourNotFound.BackColor);

            myData.ColourBold_NotFound = chkNotFoundBold.Checked;

            // ---------
            myData.TimeSheetRowNo = numTimeSheetRowNo.Value;
            myData.TimeSheetGetRowNo = chkTimeSheetGetRowNo.Checked;

            // ---------
            myData.PingSheetRowNo = numPingSheetRowNo.Value;
            myData.ColPingRead = numColPingRead.Value;
            myData.ColPingWrite = numColPingWrite.Value;

            //---- misc
            myData.TestCode = chkTestCode.Checked;
            // myData.RecordTimes = chkRecordTimes.Checked;

            myData.TurnOffScreenValidation = chkTurnOffScreenValidation.Checked;

            myData.ClearFormatting = chkClearFormatting.Checked;

            myData.FileDateTime = cmboWhichDate.Text;
            myData.ExtractFileName = chkExtractFileName.Checked;
            myData.ColExtractedFile = numColNoForExtractedName.Value;

            myData.ZapStartDefaultRow = numZapSheetStartRow.Value;

            InformationForSettingsForm.SaveData(myData);

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

        private void chkLBColourOrDelete_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked && chkLBColourOrDelete.CheckedItems.Count > 0)
            {
                chkLBColourOrDelete.ItemCheck -= chkLBColourOrDelete_ItemCheck;
                chkLBColourOrDelete.SetItemChecked(chkLBColourOrDelete.CheckedIndices[0], false);
                chkLBColourOrDelete.ItemCheck += chkLBColourOrDelete_ItemCheck;
            }

        }


        private void chkLBColourOrDelete_Leave(object sender, EventArgs e)
        {
            if (chkLBColourOrDelete.CheckedItems.Count == 0)
                chkLBColourOrDelete.SetItemChecked(0, true);

        }

        private void chkFoundBold_CheckedChanged(object sender, EventArgs e)
        {

            chkFoundBold.SwitchToBoldRegularChkBox();
            txtColourFound.SwtichToBoldRegularTextBox();

            /*
            // set font to bold
            if (chkFoundBold.Checked)
            {
                txtColourFound.Font = new Font(txtColourFound.Font, FontStyle.Bold);
                chkFoundBold.Font = new Font(chkFoundBold.Font, FontStyle.Bold);
            } else {
                txtColourFound.Font = new Font(txtColourFound.Font, FontStyle.Regular);
                chkFoundBold.Font = new Font(chkFoundBold.Font, FontStyle.Regular);

            }*/

        }

        private void chkNotFoundBold_CheckedChanged(object sender, EventArgs e)
        {
            chkNotFoundBold.SwitchToBoldRegularChkBox();
            txtColourNotFound.SwtichToBoldRegularTextBox();




            // set font to bold
            // if (chkNotFoundBold.Checked)
            // {
            //CommonExcelClasses.SwtichToBoldRegularChkBox(chkNotFoundBold);
            //CommonExcelClasses.SwtichToBoldRegularTextBox(txtColourNotFound);

            // txtColourNotFound.Font = new Font(txtColourNotFound.Font, FontStyle.Bold);
            // chkNotFoundBold.Font = new Font(chkFoundBold.Font, FontStyle.Bold);
            // }
            //  else
            // {
            // txtColourNotFound.Font = new Font(txtColourNotFound.Font, FontStyle.Regular);
            // chkNotFoundBold.Font = new Font(chkFoundBold.Font, FontStyle.Regular);

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }



    }

}
