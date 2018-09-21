using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RiggingConsoleApp
{
    class GraveYard
    {

        /*
        var result = excelReader.AsDataSet(new ExcelDataSetConfiguration() {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() {
                UseHeaderRow = true } });


                int rowPosition = 8;
                int columnPosition = 1;
                Console.WriteLine(result.Tables[0].Rows.Count);
                Console.WriteLine(result.Tables[0].Columns.Count);
                string strVal1 = result.Tables[0].Rows[rowPosition][columnPosition].ToString();
                CommonExcelClasses.MsgBox("strVal1: " + strVal1, "Information");
                strVal1 = CommonExcelClasses.myCellFormat(strVal1);
                CommonExcelClasses.MsgBox("strVal1: " + strVal1, "Information");
                // qwihtut datasedt
                Console.WriteLine(result.Tables[0].Rows[rowPosition][columnPosition]);
                // stream.Close();
                */


        //3. DataSet - Create column names from first row
        // excelReader.IsFirstRowAsColumnNames = false;


        public static double[] getCoordsFromRangeWeb_eg2(string strRange)
        {
            // var cellStr = "AB2"; // var cellStr = "A1";


            string cellStr = strRange;

            var match = Regex.Match(cellStr, @"(?<col>[A-Z]+)(?<row>\d+)");
            string colStr = match.Groups["col"].ToString();
            double col = colStr.Select((t, i) => (colStr[i] - 64) * Math.Pow(26, colStr.Length - i - 1)).Sum();
            double row = int.Parse(match.Groups["row"].ToString());

            double[] intArrReturn = { col, row };

            return intArrReturn;

        }

        public static int getCoordsFromRangeWeb_eg1(string strRange)
        {
            // var cellStr = "AB2"; // var cellStr = "A1";


            var cellStr = strRange;

            var match = Regex.Match(cellStr, @"(?<col>[A-Z]+)(?<row>\d+)");
            var colStr = match.Groups["col"].ToString();
            var col = colStr.Select((t, i) => (colStr[i] - 64) * Math.Pow(26, colStr.Length - i - 1)).Sum();
            var row = int.Parse(match.Groups["row"].ToString());

            return row;

        }


    }
}
