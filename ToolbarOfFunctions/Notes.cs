		
// time taken code			
            #region [Declare and instantiate variables for process]
            myData = myData.LoadMyData();               // read data from settings file

            bool boolDisplayInitialMessage = myData.ProduceInitialMessageBox;
            bool boolDisplayCompleteMessage = myData.ProduceCompleteMessageBox;
            bool booltimeTaken = myData.DisplayTimeTaken;
            bool boolTurnOffScreen = myData.TurnOffScreenValidation;
            bool boolClearFormatting = myData.ClearFormatting;
            bool boolTestCode = myData.TestCode;

            #endregion


                    #region [Ask to display a Message?]
                    DialogResult dlgResult = DialogResult.Yes;
                    string strMessage;
                                        
                    if (boolDisplayInitialMessage)
                    {
                        strMessage = "Compare: " + Wks1.Name + LF +
                                    " against: " + Wks2.Name + LF +
                                        " and: " + strCompareOrColour + " ones which are the same" + LF +
                                       " (starting at row:" + intStartRow.ToString() + ")";

                        if (booltimeTaken)
                        {
                            strMessage = strMessage + LF + " and display the time taken";
                        }

                        strMessage = strMessage + "?";

                        dlgResult = MessageBox.Show(strMessage, "Compare Sheets", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    }

                    if (dlgResult == DialogResult.Yes)
                    {
                        DateTime dteStart = DateTime.Now;
					
                        if (boolTurnOffScreen)
                            CommonExcelClasses.turnAppSettings("On", xls, myData.TestCode);

					
					
                        #region [Display Complete Message]
                        if (boolDisplayCompleteMessage)
                        {
                            strMessage = "";
                            strMessage = strMessage + "Compare Complete ...";

                            if (booltimeTaken)
                            {

                                DateTime dteEnd = DateTime.Now;
                                int milliSeconds = (int)((TimeSpan)(dteEnd - dteStart)).TotalMilliseconds;

                                strMessage = strMessage + "that took {TotalMilliseconds} " + milliSeconds;

                            }

                            CommonExcelClasses.MsgBox(strMessage);          // localisation?
                        }
                        #endregion


						