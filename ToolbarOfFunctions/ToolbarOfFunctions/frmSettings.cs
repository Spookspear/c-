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
        // public string strFilename = "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml";
        public string strFilename = CommonExcelClasses.strFilename;

        public frmSettings()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            // load data
            if (File.Exists(strFilename))
            {
                XmlSerializer xs = new XmlSerializer(typeof(Information));
                FileStream read = new FileStream(strFilename,FileMode.Open, FileAccess.Read, FileShare.Read);
                Information info = (Information)xs.Deserialize(read);

                chkLargeButtons.Checked = info.LargeButtons;
                chkHideText.Checked = info.HideText;
                cmboCompareDifferences.Text = info.Differences;

                cmboHighLightOrDelete.Text = info.HighLightOrDelete;
                chkDisplayTimeTaken.Checked = info.DisplayTimeTaken;
                chkProduceMessageBox.Checked = info.ProduceMessageBox;
                cmboDelModeAorBorC.Text = info.DelModeAorBorC;
                numHighlightRowsOver.Value = info.HighlightRowsOver;
                // ---------

                numNoOfColumnsToCheck.Value = info.NoOfColumnsToCheck;
                numComparingStartRow.Value = info.ComparingStartRow;
                numDupliateColumnToCheck.Value = info.DupliateColumnToCheck;
                txtColourFound.Text = info.ColourFound;
                txtColourNotFound.Text = info.ColourNotFound;

                // ---------
                numTimeSheetRowNo.Value = info.TimeSheetRowNo;
                chkTimeSheetGetRowNo.Checked = info.TimeSheetGetRowNo;

                // ---------
                numPingSheetRowNo.Value = info.PingSheetRowNo;
                numColPingRead.Value = info.ColPingRead;
                numColPingWrite.Value = info.ColPingWrite;

                read.Close();
                
            }


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


            if (cmboDelModeAorBorC.Items.Count != 3)
            {
                cmboDelModeAorBorC.Items.Add("Mode: A");
                cmboDelModeAorBorC.Items.Add("Mode: B");
                cmboDelModeAorBorC.Items.Add("Mode: C");
                // cmboDelModeAorBorC.SelectedIndex = 0;

            }


        }

        private void btnApply_Click(object sender, EventArgs e)
        {

            // now save the data
            try
            {
                Information info = new Information
                {
                    LargeButtons = chkLargeButtons.Checked,
                    HideText = chkHideText.Checked,
                    Differences = cmboCompareDifferences.Text,

                    HighLightOrDelete = cmboHighLightOrDelete.Text,
                    DisplayTimeTaken = chkDisplayTimeTaken.Checked,
                    ProduceMessageBox = chkProduceMessageBox.Checked,
                    DelModeAorBorC = cmboDelModeAorBorC.Text,
                    HighlightRowsOver = numHighlightRowsOver.Value,
                    // ---------

                    NoOfColumnsToCheck = numNoOfColumnsToCheck.Value,
                    ComparingStartRow = numComparingStartRow.Value,
                    DupliateColumnToCheck = numDupliateColumnToCheck.Value,
                    ColourFound = txtColourFound.Text,
                    ColourNotFound = txtColourNotFound.Text,

                    // ---------
                    TimeSheetRowNo = numTimeSheetRowNo.Value,
                    TimeSheetGetRowNo = chkTimeSheetGetRowNo.Checked,

                    // ---------
                    PingSheetRowNo = numPingSheetRowNo.Value,
                    ColPingRead = numColPingRead.Value,
                    ColPingWrite = numColPingWrite.Value
                };

                SaveXML.SaveData(info, "D:\\GitHub\\c-\\ToolbarOfFunctions\\ToolbarOfFunctions\\data.xml");

            }
            catch (Exception ex)
            {
                CommonExcelClasses.MsgBox(ex.Message);
            }

            // btnApply.Tag = "Apply";
            // this.Hide();
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {

            // can I dynamically change the text if a button?

            //now set the buttons size if reqd
            // boolDisplayMessage = myForm.chkLargeButtons.Checked;

            this.Hide();            
        }





    }
}
