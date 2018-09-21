#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace RiggingConsoleApp
{
    public static class CommonExcelClasses
    {

        public const int _COL1 = 0;     // A
        public const int _ROW1 = 1;     // 6

        public const int _COL2 = 2;
        public const int _ROW2 = 3;

        public static void MsgBox(string strMessage, string strWhichIcon = "Information")
        {
            MessageBoxIcon whichIcon = MessageBoxIcon.Information;
            string strCaption = strWhichIcon;

            switch (strWhichIcon)
            {
                case "Question":
                    whichIcon = MessageBoxIcon.Question;
                    break;

                case "Error":
                    whichIcon = MessageBoxIcon.Error;
                    break;

                case "Information":
                    whichIcon = MessageBoxIcon.Information;
                    break;

                case "Warning":
                    whichIcon = MessageBoxIcon.Warning;
                    break;

            }

            MessageBox.Show(strMessage, strCaption, MessageBoxButtons.OK, whichIcon);
        }

        public static bool checkEmptyRange(DataTable Wks, string strRange, int intOffset)
        {
            double[] dblArrCoords = getCoordsFromRange1(strRange);
            int intNotEmpty = 0;
            int iRow;
            int iCol;

            MsgBox("Not tested");

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

        public static string myCellFormat(string strCell)
        {
            strCell = strCell.Replace("\r", " ");
            strCell = strCell.Replace("\n", " ");
            strCell = strCell.Replace("\t", " ");
            strCell = strCell.Replace("  ", " ");
            return strCell;
        }

        internal static int searchForValue(DataTable dataTable, string searchString, int intScanCol)
        {

            // might have to do this the old way - loop both dimentions
            // var dataTable = wksNew.Tables[0];

            int intRow = 1;
            int intRetVal = 0;

            for (int iRows = intRow; iRows < dataTable.Rows.Count; iRows++)
            {

                Console.WriteLine(dataTable.Rows[iRows][intScanCol].ToString());

                if (dataTable.Rows[iRows][intScanCol].ToString() == searchString)
                {
                    intRetVal = iRows;
                    break;
                }
                /*
                if (getDBIntValue(dataTable.Rows[iRows][intScanCol], 0) > 0)
                {
                    if ((string)dataTable.Rows[iRows][intScanCol].ToString() == searchString)
                    {
                        intRetVal = iRows;
                        break;
                    }
                }
                */

            }

            return intRetVal;

        }


        public static int getDBIntValue(object value, int defaultValue)
        {
            if (!Convert.IsDBNull(value))
            {
                return (int)value;
            }
            else
            {
                return defaultValue;
            }
        }


        public static double[] getCoordsFromRange1(string strRange)
        {
            // var cellStr = "AB2"; // var cellStr = "A1";
            double dblCol1 = 0; double dblRow1 = 0; double dblCol2 = 0; double dblRow2 = 0;
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
            dblCol1 = CommonExcelClasses.getExcelColumnNumber(strAddress.Substring(0, intIdx));
            dblRow1 = Convert.ToInt32(strAddress.Substring(intIdx));

            if (intSPos > 0)
            {
                strAddress = strRange.Substring(intSPos + 1);
                intIdx = strAddress.IndexOfAny("0123456789".ToCharArray());
                dblCol2 = CommonExcelClasses.getExcelColumnNumber(strAddress.Substring(0, intIdx));
                dblRow2 = Convert.ToInt32(strAddress.Substring(intIdx));
            }

            double[] intArrReturn = { dblCol1, dblRow1, dblCol2, dblRow2 };

            return intArrReturn;

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

        public static DateTime getFileDate(string strFileName)
        {
            FileInfo oFileInfo = new FileInfo(strFileName);
            FileVersionInfo oFileVersionInfo = FileVersionInfo.GetVersionInfo(strFileName);

            return oFileInfo.LastWriteTime;

        }

    }

}
