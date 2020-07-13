#pragma warning disable IDE1006 // Naming Styles

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ToolbarOfFunctions_CommonClasses;

using System.Windows.Forms;                 // for ok prompt
using Microsoft.Office.Interop.Excel;

namespace ToolbarOfFunctions
{
	public partial class ThisAddIn
	{
		readonly Excel.XlSheetVisibility xlSheetUnHide = Excel.XlSheetVisibility.xlSheetVisible;
		readonly Excel.XlSheetVisibility xlSheetHide = Excel.XlSheetVisibility.xlSheetHidden;
		readonly string strAllSheetCodeNames = "WksActions;WksBuying;WksCompPosAnalysis;WksCompPosChart;WksParameters;WksProspectInfo;WksValuePropositions;WksWarningSheet";

		public void toggleWorksheetVisability( Excel.Application xls )
		{
			Excel.Workbook Wkb = xls.ActiveWorkbook;

			DialogResult dlgResult = DialogResult.Yes;

			string strMessage;

			// declare array of known names - not needed
			// loop through sheets looking at names
			strMessage = "";
			strMessage += "Workbook: {0} hide or unhide required worksheets" + LF + LF;
			strMessage += "Yes = Hide worksheets" + LF;
			strMessage += "No = UnHide worksheets" + LF;
			strMessage += "Cancel = exit routine" + LF;
			strMessage += "?";
			strMessage = string.Format(strMessage, Wkb.CodeName);
			dlgResult = MessageBox.Show(strMessage, "Question", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

			if ( dlgResult != DialogResult.Cancel ) {
				if ( dlgResult == DialogResult.Yes ) {
					toggleVisability(Wkb, "Hide");
				}

				if ( dlgResult == DialogResult.No ) {
					toggleVisability(Wkb, "Unhide");
				}

			}
		}

		/// <summary>
		/// Will toggle the Visability of known worksheets
		/// Should maybe get the names from the workbook itself
		/// </summary>
		/// <param name="Wkb"></param>
		/// <param name="v"></param>
		private void toggleVisability( Excel.Workbook Wkb, string v )
		{
			// turn off screen updating
			turnAppSettingsNew("Off");

			string[] arrWorkSheetCodeName = strAllSheetCodeNames.Split(';');
			Excel.Worksheet WksWarningSheet = Wkb.Sheets["CRM365"];
			Excel.Worksheet WksProspectInfo = Wkb.Sheets["Prospect Info"];

			// as were hiding all sheets need at least one visible
			if ( v == "Hide" ) {
				WksWarningSheet.Visible = xlSheetUnHide;
			}

			// loop through all sheets 
			foreach ( Excel.Worksheet Wks in Wkb.Sheets ) {

				// check worksheets is one of allowed sheets
				if ( aScan(Wks.CodeName, arrWorkSheetCodeName) ) {

					if ( v == "Hide" ) {
						if ( Wks.CodeName != "WksWarningSheet" )
							if ( Wks.Visible != xlSheetHide )
								Wks.Visible = xlSheetHide;
					}

					if ( v == "Unhide" ) {
						if ( Wks.CodeName != "WksWarningSheet" && Wks.CodeName != "WksParameters" )
							if ( Wks.Visible != xlSheetUnHide )
								Wks.Visible = xlSheetUnHide;

					}
				}

			}

			if ( v == "Unhide" ) {
				WksWarningSheet.Visible = xlSheetHide;
				WksProspectInfo.Select();
			}


			turnAppSettingsNew("On");
		}

		private bool aScan( string codeName, string [] arrWorkSheetCodeName )
		{
			bool bRetVal = false;

			for ( int i = 0; i < arrWorkSheetCodeName.GetLength(0); i++ ) {
				if ( arrWorkSheetCodeName [i].ToLower() == codeName.ToLower() ) {
					bRetVal = true;
					break;
				}
			}
			return bRetVal;
		}

		private static void turnAppSettingsNew( string strDoWhat )
		{
			bool boolOn = true;
			Excel.Application xls = Globals.ThisAddIn.Application;

			if ( strDoWhat == "Off" )
				boolOn = false;

			xls.EnableEvents = boolOn;
			xls.ScreenUpdating = boolOn;

			if ( boolOn ) {
				xls.Cursor = Excel.XlMousePointer.xlDefault;
				xls.Calculation = XlCalculation.xlCalculationAutomatic;
			} else {
				xls.Cursor = Excel.XlMousePointer.xlWait;
				xls.Calculation = XlCalculation.xlCalculationManual;
			}

		}

		public void InsertNewWorksheet(Excel.Application xls)
		{
			Excel.Worksheet newWorksheet;
			newWorksheet = ( Excel.Worksheet )this.Application.Worksheets.Add();

			var newSheet = (Microsoft.Office.Interop.Excel.Worksheet)xls.Worksheets.Add(Type.Missing, xls.Worksheets[xls.Worksheets.Count], 1, XlSheetType.xlWorksheet);
			newSheet.Name = "myWorkSheet";

			// trying to chnage the sheets code name - doesnt appear to be possible
			var newSheet2 = xls.Worksheets.Add(Type.Missing, xls.Worksheets[xls.Worksheets.Count], 1, XlSheetType.xlWorksheet) as Worksheet;
			// newSheet2.CodeName = "Wks";

		}

	}

}
