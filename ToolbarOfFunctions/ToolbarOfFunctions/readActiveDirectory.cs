#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

using System.IO;            // for Directory function
using System.Diagnostics;   // .FileVersionInfo
using System.Drawing;       // for colours

using DaveChambers.FolderBrowserDialogEx;

using System.ComponentModel;
using System.Data;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Tools.Ribbon;

using ToolbarOfFunctions_CommonClasses;
using ToolbarOfFunctions_MyConstants;
using System.Runtime.InteropServices;

// using System.Data.SqlTypes;

using System.DirectoryServices;


namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {
   
        public void loadADGroupIntoSpreadsheetActiveCell(Excel.Application xls)
        {
            CommonExcelClasses.MsgBox("loadADGroupIntoSpreadsheetActiveCell - is what this will do");


        }

    }

}
