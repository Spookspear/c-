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

using OfficeOpenXml;

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

        public static bool checkEmptyRange(ExcelWorksheet Wks, string strRange, int i)
        {
            int[] intArrCoords = getCoordsFromRange(strRange);
            int intNotEmpty = 0;
            
            for (int x = intArrCoords[_ROW1]; x <= intArrCoords[_ROW2]; x++)                // loop along the Rows
            {                
                for (int y = intArrCoords[_COL1]; y <= intArrCoords[_COL2]; y++)            // loop along the Cols
                {
                    if (!isEmptyCell(Wks.Cells[(x - i), (y - i)]))
                    {
                        if (Wks.Cells[(x - i), (y - i)].Value.ToString() != "")
                        {
                            intNotEmpty++;
                            break;
                        }
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

        internal static int searchForValue(ExcelWorksheet wks, string searchString, int intScanCol)
        {
            int intRow = 1;
            int intRetVal = 0;

            for (int iRows = intRow; iRows < wks.Dimension.End.Row; iRows++)
            {
                if (!isEmptyCell(wks.Cells[iRows, intScanCol]))
                {
                    Console.WriteLine(wks.Cells[iRows, intScanCol].Value.ToString());

                    if (wks.Cells[iRows, intScanCol].Value.ToString() == searchString)
                    {
                        intRetVal = iRows;
                        break;
                    }
                }

            }

            return intRetVal;

        }

        public static bool isEmptyCell(ExcelRange xlCell)
        {
            bool boolRetVal = false;

            if (xlCell.Value == null)
            {
                boolRetVal = true;
            }

            return boolRetVal;

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


        public static int[] getCoordsFromRange(string strRange)
        {
            int intCol1 = 0, intRow1 = 0, intCol2 = 0, intRow2 = 0;
            int intSPos;
            string strAddress1, strAddress2;

            intSPos = strRange.IndexOf(":");

            if (intSPos < 0)
                strAddress1 = strRange;
            else
                strAddress1 = strRange.Substring(0, intSPos);

            // stick into return values
            intCol1 = strAddress1.Col();
            intRow1 = strAddress1.Row();

            if (intSPos > 0)
            {
                strAddress2 = strRange.Substring(intSPos + 1);
                intCol2 = strAddress2.Col();
                intRow2 = strAddress2.Row();
            }

            int[] intArrReturn = { intCol1, intRow1, intCol2, intRow2 };

            return intArrReturn;

        }

        public static DateTime getFileDate(string strFileName)
        {
            FileInfo oFileInfo = new FileInfo(strFileName);
            FileVersionInfo oFileVersionInfo = FileVersionInfo.GetVersionInfo(strFileName);

            return oFileInfo.LastWriteTime;

        }


        // extensability - will it work?


        public static int Col(this string strRange)
        {
            var match = Regex.Match(strRange, @"(?<col>[A-Z]+)(?<row>\d+)");
            string colStr = match.Groups["col"].ToString();
            double col = colStr.Select((t, i) => (colStr[i] - 64) * Math.Pow(26, colStr.Length - i - 1)).Sum();
            return Convert.ToInt32(col);

        }

        public static int Row(this string strRange)
        {
            var match = Regex.Match(strRange, @"(?<col>[A-Z]+)(?<row>\d+)");
            string colStr = match.Groups["col"].ToString();
            double row = int.Parse(match.Groups["row"].ToString());
            return Convert.ToInt32(row);

        }



        public static string Message(this string strMessage)
        {

            string strWhichIcon = "Information";
            MessageBoxIcon whichIcon = MessageBoxIcon.Information;
            string strCaption = strWhichIcon;

            MessageBox.Show(strMessage, "Information", MessageBoxButtons.OK, whichIcon);

            return strMessage;

        }



    }

}
