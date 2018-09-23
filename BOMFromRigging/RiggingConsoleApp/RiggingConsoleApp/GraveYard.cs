using System;
using System.Collections.Generic;
using System.Data;
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

        public const int _COL1 = 0;     // A
        public const int _ROW1 = 1;     // 6

        public const int _COL2 = 2;
        public const int _ROW2 = 3;


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


        public static double[] getCoordsFromRange1(string strRange)
        {
            // var cellStr = "AB2"; // var cellStr = "A1";
            double dblCol1 = 0, dblRow1 = 0, dblCol2 = 0, dblRow2 = 0;
            int intSPos;
            string strAddress;

            intSPos = strRange.IndexOf(":");

            if (intSPos < 0)
                strAddress = strRange;
            else
                strAddress = strRange.Substring(0, intSPos);


            double[] dblArrColRow = getArrGetColRow(strAddress);

            // stick into return values
            dblCol1 = dblArrColRow[0];
            dblRow1 = dblArrColRow[1];

            if (intSPos > 0)
            {
                strAddress = strRange.Substring(intSPos + 1);
                dblArrColRow = getArrGetColRow(strAddress);

                // stick into return values
                dblCol2 = dblArrColRow[0];
                dblRow2 = dblArrColRow[1];

            }

            double[] intArrReturn = { dblCol1, dblRow1, dblCol2, dblRow2 };

            return intArrReturn;

        }

        private static double[] getArrGetColRow(string cellStr)
        {
            var match = Regex.Match(cellStr, @"(?<col>[A-Z]+)(?<row>\d+)");
            string colStr = match.Groups["col"].ToString();
            double col = colStr.Select((t, i) => (colStr[i] - 64) * Math.Pow(26, colStr.Length - i - 1)).Sum();
            double row = int.Parse(match.Groups["row"].ToString());

            double[] dblArrReturn = { col, row };

            return dblArrReturn;

        }



        public static int getExcelColumnNumber(string strLetter)
        {
            strLetter = strLetter.ToUpper();
            int intOutNum = 0;

            for (int i = 0; i < strLetter.Length; i++)
            {
                intOutNum *= 26;
                intOutNum += (strLetter[i] - 'A' + 1);
            }
            return intOutNum;

        }


        public static double[] getCoordsFromRange2(string strRange)
        {

            double dblCol1 = 0; double dblRow1 = 0; double dblCol2 = 0; double dblRow2 = 0;
            string strAddress;
            int intSPos;

            intSPos = strRange.IndexOf(":");

            if (intSPos < 0)
                strAddress = strRange;
            else
                strAddress = strRange.Substring(0, intSPos);


            // A20:F20 but could be AA20:FF20 - so now need to get 1st occurance of a number
            int intIdx = strAddress.IndexOfAny("0123456789".ToCharArray());
            dblCol1 = getExcelColumnNumber(strAddress.Substring(0, intIdx));
            dblRow1 = Convert.ToInt32(strAddress.Substring(intIdx));

            if (intSPos > 0)
            {
                strAddress = strRange.Substring(intSPos + 1);
                intIdx = strAddress.IndexOfAny("0123456789".ToCharArray());
                dblCol2 = getExcelColumnNumber(strAddress.Substring(0, intIdx));
                dblRow2 = Convert.ToInt32(strAddress.Substring(intIdx));
            }

            double[] intArrReturn = { dblCol1, dblRow1, dblCol2, dblRow2 };

            return intArrReturn;

        }

        public static bool checkEmptyRangeOld(DataTable Wks, string strRange, int intOffset)
        {

            double[] dblArrCoords = getCoordsFromRange1(strRange);
            int intNotEmpty = 0;
            int iRow;
            int iCol;


            // loop along the Rows
            for (double x = dblArrCoords[_ROW1]; x <= dblArrCoords[_ROW2]; x++)
            {
                // loop along the Cols
                for (double y = dblArrCoords[_COL1]; y <= dblArrCoords[_COL2]; y++)
                {
                    iRow = Convert.ToInt32(x) - intOffset;
                    iCol = Convert.ToInt32(y) - intOffset;

                    if (Wks.Rows[iRow][iCol].ToString() == "")
                    {
                        intNotEmpty++;
                        break;
                    }
                }

                if (intNotEmpty > 0)
                    break;

            }

            return (intNotEmpty == 0);

        }

        private static string getExcelValueOld(DataTable dataTable, string strAddress, int intOffset)
        {
            double[] dblArrCoords = getCoordsFromRange1(strAddress);

            int iCol = Convert.ToInt32(dblArrCoords[_COL1]);
            int iRow = Convert.ToInt32(dblArrCoords[_ROW1]);

            string strRetVal = dataTable.Rows[(iRow - intOffset)][(iCol - intOffset)].ToString();

            return strRetVal;


        }

        private static void TestCode()
        {
            //string strRange = "AB1";

            string strRange;
            string strMessage;

            const int _ROW1 = 0;
            const int _COL1 = 1;

            const int _ROW2 = 2;
            const int _COL2 = 3;



            double[] intArrCoords;


            // 1 dimension
            strRange = "AB15";

            // web way

            intArrCoords = getCoordsFromRange1(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");


            // my way
            intArrCoords = getCoordsFromRange2(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");



            // 2 dimension
            strRange = "AB15:AZ15";

            intArrCoords = getCoordsFromRange1(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");


            // my way
            strRange = "AB15:AZ15";

            intArrCoords = getCoordsFromRange2(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");




        }




    }


}
