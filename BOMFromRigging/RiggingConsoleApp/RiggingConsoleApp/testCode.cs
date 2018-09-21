using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RiggingConsoleApp
{
    class testCode
    {

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

            intArrCoords = CommonExcelClasses.getCoordsFromRange1(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");


            // my way
            intArrCoords = CommonExcelClasses.getCoordsFromRange2(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");



            // 2 dimension
            strRange = "AB15:AZ15";

            intArrCoords = CommonExcelClasses.getCoordsFromRange1(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");


            // my way
            strRange = "AB15:AZ15";

            intArrCoords = CommonExcelClasses.getCoordsFromRange2(strRange);

            // string interpolation
            strMessage = string.Format("strRange= {0} is: Row1: {1} and Col1: {2}, Row2: {3}, Col2: {4} ", strRange, intArrCoords[_ROW1].ToString(), intArrCoords[_COL1].ToString(), intArrCoords[_ROW2].ToString(), intArrCoords[_COL2].ToString());

            CommonExcelClasses.MsgBox(strMessage, "Information");




        }


    }
}
