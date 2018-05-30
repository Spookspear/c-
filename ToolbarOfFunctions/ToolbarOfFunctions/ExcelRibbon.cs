#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ToolbarOfFunctions = ToolbarOfFunctions.ThisAddIn;


namespace ToolbarOfFunctions
{
    public partial class ExcelRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void btnZap_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.zapWorksheet(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void btnReadFolders_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.readFolders(Globals.ThisAddIn.Application.ActiveWorkbook);
        }       

        private void btnCompareSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.compareSheets(Globals.ThisAddIn.Application.ActiveWorkbook);
        }
        
        private void btnDeleteBlankLinesA_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "A");
        }


        private void btnDeleteBlankLinesB_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "B");

        }

        private void btnDeleteBlankLinesC_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook, "C");

        }

        private void splitButtonDeleteLines_Click(object sender, RibbonControlEventArgs e)
        {
            btnDeleteBlankLinesB_Click(sender, e);
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.openSettingsForm(Globals.ThisAddIn.Application.ActiveWorkbook);

            btnDealWithSingleDuplicates.Label = "Hi";


        }

        public void btnDealWithSingleDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.dealWithSingleDuplicates(Globals.ThisAddIn.Application.ActiveWorkbook);

        }
    }

}
