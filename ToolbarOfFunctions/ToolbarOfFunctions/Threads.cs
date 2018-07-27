#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using ToolbarOfFunctions_CommonClasses;

using System.Threading;

namespace ToolbarOfFunctions
{
    public partial class ThisAddIn
    {
        public void testThreads1(Excel.Application xls, string strDoWhat)
        {
            CommonExcelClasses.MsgBox("testThreads1 - break now");

        }

    }

}

