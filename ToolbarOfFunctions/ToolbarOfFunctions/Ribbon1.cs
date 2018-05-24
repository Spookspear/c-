#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ToolbarOfFunctions
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnZap_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.zapWorksheet(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void btnDeleteBlankLines_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.deleteBlankLines(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        private void btnReadFolders_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.readFolders(Globals.ThisAddIn.Application.ActiveWorkbook);
        }       

        private void btnCompareSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.compareSheets(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

    }

}
