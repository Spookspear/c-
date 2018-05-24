Attribute VB_Name = "btnSuttonCallSubs"
Option Explicit

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnPopulateSheetFromFolder
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' Returns   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnPopulateSheetFromFolder()
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes
Dim intReadDetails As Integer: intReadDetails = vbYes
Dim intHighlightRows As Integer

    If notParameters(ActiveSheet) Then
    
        strAskUser = getRangeValue("rangeProduceMessageBox")
        
        If strAskUser = "Y" Then
            intProceed = MsgBox("Read 'Folder/Directory' into Worksheet: " & ActiveSheet.Name & " ?", vbQuestion + vbYesNo)
            
            If intProceed = vbYes Then
                intReadDetails = MsgBox("Populate details column in " & ActiveSheet.Name & " ?", vbQuestion + vbYesNo)
            End If
            
        End If
        
        If intProceed = vbYes Then
        
            intHighlightRows = getRangeValue("rangeHighlightRows")
        
            If populateSheetFromFolder(ActiveSheet.Name, intHighlightRows) Then
                Application.ScreenUpdating = False
                If intReadDetails = vbYes Then
                    Call grabDetails(ActiveSheet.Name)
                End If
                
                Call writeHeaders(ActiveSheet, "FILES")
                ' Set the number of compare columns to 5 - as we are probably going to compare sheets
                Call setRangeValue("rangeNoOfColumnsToCheck", 5)
                
            End If
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnCompareSheets
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' Returns   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnCompareSheets()
Dim strSheetName1 As String
Dim strSheetName2 As String
Dim intActiveSheetIndex As Integer
Dim strClearOrColour As String
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
        ' Whatever sheet is selected compare it against the one next door
        intActiveSheetIndex = ActiveSheet.Index
        
        strSheetName1 = Sheets(intActiveSheetIndex).Name
        strSheetName2 = Sheets(intActiveSheetIndex + 1).Name
    
        strClearOrColour = getRangeValue("rangeCompareOption")
        strAskUser = getRangeValue("rangeProduceMessageBox")
    
        If strAskUser = "Y" Then
            intProceed = MsgBox("Compare: Worksheet: " & strSheetName1 & " against: " & strSheetName2 & " and " & strClearOrColour & " ones which are the same?", vbQuestion + vbYesNo)
        End If
        
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            Call compareSheets(Worksheets(strSheetName1), Worksheets(strSheetName2))
        End If
        
    End If
    
End Sub
Sub Test()
    ' how do i get the active sheet no?
    

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnZapSheet
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnZapSheet()
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    strAskUser = getRangeValue("rangeProduceMessageBox")

    If notParameters(ActiveSheet) Then
        
        If strAskUser = "Y" Then
            intProceed = MsgBox("ZAP: Worksheet {" & ActiveSheet.Name & "} ?", vbQuestion + vbYesNo)
        End If

        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            Call ZapWorkSheet(ActiveSheet, 1)
            'Call writeHeaders(ActiveSheet, "ZAP")
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnDelBlankLines
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   | updated to include timer code - to see how fast on Windows 7
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnDelBlankLines()
Dim strStartTime, strEndTime, strTotalTime
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
    
        strAskUser = getRangeValue("rangeProduceMessageBox")
        
        If strAskUser = "Y" Then
            intProceed = MsgBox("DELETE: Blank Lines on {" & ActiveSheet.Name & "} ? using Mode: " & getRangeValue("rangeDelBlankLinesModeAorB"), vbQuestion + vbYesNo)
        End If
            
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            strStartTime = Time
            
            Call delBlankLines(ActiveSheet, getRangeValue("rangeDelBlankLinesModeAorB"))
            
            strEndTime = Time
            strTotalTime = (strEndTime - strStartTime)
        
            If getRangeValue("rangeTimeTaken") Then
                MsgBox "Time taken: " & Format(strTotalTime, "hh:mm:ss")
            End If
            
        End If
        
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnDealWithSingleDuplicates
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnDealWithSingleDuplicates()
Dim strHighlighOrDeleteDupliates As String
Dim intColumToCheck As Integer
Dim strColumToCheckName As String
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    strAskUser = getRangeValue("rangeProduceMessageBox")

    If notParameters(ActiveSheet) Then
        strHighlighOrDeleteDupliates = getRangeValue("rangeHighlightOrDeleteOption")
        intColumToCheck = getRangeValue("rangeDupliateColumnToCheck")
        
        With ActiveSheet
            strColumToCheckName = .Cells(1, intColumToCheck).Value
        End With
        
        If strAskUser = "Y" Then
            intProceed = MsgBox(UCase(strHighlighOrDeleteDupliates) & " Duplicate Rows Check on column {" & getColLetter(intColumToCheck) & "} worksheet name: {" & ActiveSheet.Name & "} ?", vbQuestion + vbYesNo)
        End If
        
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            Call dealWithSingleDuplicates(ActiveSheet)
            Call NavigateHome(ActiveSheet)
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnDealWithManyDuplicates
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnDealWithManyDuplicates()
Dim strHighlighOrDeleteDupliates As String
Dim intNoCheckCols As Integer
Dim intColumToCheck As Integer
Dim strColumToCheckName As String
Dim strMessage As String
Dim curNoDuplicates As Currency
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
    
        strAskUser = getRangeValue("rangeProduceMessageBox")
        
        strHighlighOrDeleteDupliates = getRangeValue("rangeHighlightOrDeleteOption")
        intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck")
        intColumToCheck = getRangeValue("rangeDupliateColumnToCheck")
        
        With ActiveSheet
            strColumToCheckName = .Cells(1, intColumToCheck).Value
        End With
        
        strMessage = strColumToCheckName & " ?"
        
        strMessage = ""
        strMessage = strMessage & "Worksheet name: '" & ActiveSheet.Name & "'" & vbNewLine
        strMessage = strMessage & UCase(strHighlighOrDeleteDupliates) & ": Starting Row: " & getRangeValue("rangeComparingStartRow") & vbNewLine
        strMessage = strMessage & "Starting: " & getColLetter(intColumToCheck) & " ('" & strColumToCheckName & "') for "
        strMessage = strMessage & CStr(intNoCheckCols) & " Columns" & vbNewLine
        
        If strAskUser = "Y" Then
            intProceed = MsgBox(strMessage, vbQuestion + vbYesNo)
        End If
        
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            curNoDuplicates = dealWithManyDuplicates(ActiveSheet, intColumToCheck)
            Call NavigateHome(ActiveSheet)
            
            If getRangeValue("RangeProduceMessageBox") = "Y" Then
                If curNoDuplicates > 0 Then
                    MsgBox "Handled: " & CStr(curNoDuplicates) & " duplicates", vbInformation
                End If
                
            End If
            
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnLoadADGroupIntoSpreadsheet
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnLoadADGroupIntoSpreadsheet()
Dim strActiveSheetName As String
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    strActiveSheetName = ActiveSheet.Name
   
    If notParameters(ActiveSheet) Then
    
        strAskUser = getRangeValue("rangeProduceMessageBox")
        
        If strActiveSheetName = "@" Then
            strActiveSheetName = InputBox(" Long name detected - enter here: {s7@groupname}", "Long name detected - enter here:")
            If Len(Trim(strActiveSheetName)) > 0 Then intProceed = vbYes
        Else
            If strAskUser = "Y" Then
                intProceed = MsgBox("Retrieve members from AD Group: " & strActiveSheetName & " into active worksheet?", vbQuestion + vbYesNo)
            End If
        End If
        
        If intProceed Then
            Application.ScreenUpdating = False
            Call LoadADGroupIntoSpreadsheet(ActiveSheet, GrabDomainOrGroup(strActiveSheetName, "Domain", "@"), GrabDomainOrGroup(strActiveSheetName, "User", "@"))
            Call writeHeaders(ActiveSheet, "GROUP")
            
            ' now rename sheet if its called: @ to what was determined
            ActiveSheet.Name = Mid(GrabDomainOrGroup(strActiveSheetName, "User", "@"), 1, 31)
            
            'MsgBox "bREAK NOW"
            'Call SortSheet(ActiveSheet)
            
        End If
        
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnLoadADGroupIntoSpreadsheetActiveCell
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnLoadADGroupIntoSpreadsheetActiveCell()
Dim strRememberValue As String
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
        strAskUser = getRangeValue("rangeProduceMessageBox")
        strRememberValue = ActiveCell.Value
        
        If Len(Trim(strRememberValue)) > 0 Then
            If strAskUser = "Y" Then
                intProceed = MsgBox("Retrieve members from AD Group: " & vbNewLine & vbNewLine & ActiveCell.Value & vbNewLine & vbNewLine & "into a new worksheet?", vbQuestion + vbYesNo)
            End If
        
            If intProceed = vbYes Then
                Application.ScreenUpdating = False
                Sheets.Add After:=Sheets(Sheets.Count)
                ActiveSheet.Name = Mid(strRememberValue, 1, 31)
                Call LoadADGroupIntoSpreadsheet(ActiveSheet, "S7", strRememberValue)
                Call writeHeaders(ActiveSheet, "GROUP")
            End If
            
        Else
            MsgBox "No data in selection"
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | GrabDomainOrGroup
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Function GrabDomainOrGroup(strPassInValue As String, strGrabWhat As String, strSplitChar As String) As String
Dim strReturnValue As String
    Select Case strGrabWhat
    Case "Domain"
        If InStr(strPassInValue, strSplitChar) > 0 Then
            strReturnValue = Mid(strPassInValue, 1, InStr(strPassInValue, strSplitChar) - 1)
        Else
            'MsgBox "No Domain specified - using S7 (use @ to split domain from string)", vbCritical
            strReturnValue = "S7"
        End If
    Case "User"
        If InStr(strPassInValue, strSplitChar) > 0 Then
            strReturnValue = Mid(strPassInValue, InStr(strPassInValue, strSplitChar) + 1, Len(strPassInValue))
        Else
            strReturnValue = strPassInValue
        End If
    End Select
    GrabDomainOrGroup = strReturnValue
End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnReadUsersGroupMembership
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnReadUsersGroupMembership()
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
        strAskUser = getRangeValue("rangeProduceMessageBox")
    
        If strAskUser = "Y" Then
            intProceed = MsgBox("IMPORTANT: Populate Worksheet with USER NAME := " & ActiveSheet.Name & "   Group Membership?", vbQuestion + vbYesNo)
        End If
    
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            Call ZapWorkSheet(ActiveSheet, 1)
            Call readGroupMembership(ActiveSheet)
            Call writeHeaders(ActiveSheet, "USERS")
            Call sortSheet(ActiveSheet)
        End If
    End If
End Sub


' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnReadUsers
' Author    | G Bishop
' Date      | 14-05-2018
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 14-05-2018 | caller for reading users in column
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnReadUsers()
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes

    If notParameters(ActiveSheet) Then
        strAskUser = getRangeValue("rangeProduceMessageBox")
    
        If strAskUser = "Y" Then
            intProceed = MsgBox("IMPORTANT: Populate Worksheet with USER NAME := " & ActiveSheet.Name & "   Group Membership?", vbQuestion + vbYesNo)
        End If
    
        If intProceed = vbYes Then
            Application.ScreenUpdating = False
            Call ZapWorkSheet(ActiveSheet, 1)
            Call readADUserDetails(ActiveSheet)
            Call writeHeaders(ActiveSheet, "USERS")
            Call sortSheet(ActiveSheet)
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnClearTimeSheet
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnClearTimeSheet()
    If notParameters(ActiveSheet) Then
        If MsgBox("Colour in this sheet?", vbQuestion + vbYesNo) = vbYes Then
            'Application.ScreenUpdating = False
            Call timeSheetSnapIn_ColorCode

        End If
    End If
End Sub


' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnClearTimeSheet
' Author    | G Bishop
' Date      | 08/07/2016
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 08/07/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnWriteTimeSheet()
Dim intRowNo As Integer
Dim strAskUser As String

    If notParameters(ActiveSheet) Then
    
        intRowNo = getRangeValue("rangeTimeSheetRowNo")
    
        If MsgBox("Update hours in this worksheet? (starting from row number: " & intRowNo & ")", vbQuestion + vbYesNo) = vbYes Then
            'Application.ScreenUpdating = False
            Call timeSheetSnapIn_writeTimeSheet

            strAskUser = getRangeValue("rangeProduceMessageBox")
            
            If strAskUser = "Y" Then
                MsgBox "Complete ...", vbInformation
            End If
    
        End If
    End If
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | btnPingServers
' Author    | G Bishop
' Date      | 01-05-2018
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Sub btnPingServers()
Dim intRowNo As Integer
Dim strAskUser As String
Dim intPingColRead As Integer
Dim intPingColWrite As Integer

    If notParameters(ActiveSheet) Then
    
        intRowNo = getRangeValue("rangePingSheetRowNo")
        
        intPingColRead = getRangeValue("rangeColPingRead")
        intPingColWrite = getRangeValue("rangeColPingWrite")
    
        If MsgBox("Ping all servers in column: " & getColLetter(intPingColRead) & "  (starting from row number: " & intRowNo & ")", vbQuestion + vbYesNo) = vbYes Then
            'Application.ScreenUpdating = False
            'Call pingServers
            Call pingProcessSheet(ActiveSheet)
            strAskUser = getRangeValue("rangeProduceMessageBox")
            
            If strAskUser = "Y" Then
                MsgBox "Complete ...", vbInformation
            End If
    
        End If
    End If
End Sub



'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
'Sub btnColourInTimeSheet()
'    If notParameters(ActiveSheet) Then
'        If MsgBox("TEST: IMPORTANT: Colour in this sheet?", vbQuestion + vbYesNo) = vbYes Then
'            Application.ScreenUpdating = False
'            Call TimeSheetSnapIn_ColorCode
'
'        End If
'    End If
'End Sub


