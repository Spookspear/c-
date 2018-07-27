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

        public void testThreads(Excel.Application xls)
        {
            CommonExcelClasses.MsgBox("testThreads1 - break now");
            //int intSourceRow2 = 2;
            // get worksheet name
            Excel.Workbook Wkb = xls.ActiveWorkbook;
            Excel.Worksheet Wks;   // get current sheet
            Wks = Wkb.ActiveSheet;


            for (int i = 0; i < 10; i++)
            {
                ThreadedWorker tr = new ThreadedWorker(i);

                
            }

        }


        public class ThreadedWorker
        {

            int ID;
            Thread t;


            public ThreadedWorker(int ID)
            {
                this.ID = ID;
                t = new Thread(new ThreadStart(doWork));
                t.Start();
            }

            void doWork()
            {
                for (int i = 0; i < 10; i++)
                {
                    Console.WriteLine("Thread " + ID + " is running");
                }
                Console.WriteLine("Thread" + ID + " is finished" );

            }

        }

    }
}


