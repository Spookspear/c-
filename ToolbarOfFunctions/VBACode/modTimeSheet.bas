Attribute VB_Name = "modTimeSheet"
Option Explicit
' --------------------------------------------------------------------------------------
'    Author | Grant Bishop
'      Date | 08/05/2018
'      Name | modTimeSheet
'   Purpose | Holds all Timesheet code / tools
'     To Do |
' ----------+---------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Global Const DBLFORECOLOR = 1
Global Const DBLBACKCOLOR = 2
' Dim G_boolDevelopmentMode As Boolean

Const C_COL_CATEGORY = 17
Const C_COL_TOTAL = 19
Const C_COL_DATE = 14

' --------------------------------------------------------------------------------------
' Procedure | addTimeSheet
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub addTimeSheet(ByRef cbMainMenuBar As CommandBar, ByRef intButtonCount, intStyle As Currency, boolInclude As Boolean)
Dim myNewButton As CommandBarButton

    If boolInclude Then
    
        '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        ' Mypersonal timesheet option - remove if published
        '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        intButtonCount = intButtonCount + 1
        Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
        With myNewButton
            .Caption = "&Update timesheet"
            .TooltipText = "will clear out my timesheet"
            .Style = intStyle
            .FaceId = 704
            .OnAction = "btnWriteTimeSheet"
        End With
    
    End If

End Sub

' --------------------------------------------------------------------------------------
' Procedure | addValidationToColumn
' Author    | Grant Bishop
' Date      | 25/01/2016
' Purpose   | adds validation to all cells (intStartRow->endrow) in a certain column
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub addValidationToColumn(Wks As Worksheet, strCol As String, intStartRow As Integer, strEndRow As Integer, strFormula As String)
Dim strSelection As String

    strSelection = strCol & intStartRow & ":" & strCol & strEndRow
    With Wks.Range(strSelection)

        If G_boolDevelopmentMode Then
            .Activate
        End If
        
        .Locked = False

        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=strFormula
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = "Please choose a value from dropdown list."
            .ShowInput = True
            .ShowError = True
        End With

    End With

End Sub

' --------------------------------------------------------------------------------------
' Procedure | colorRow
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     | maybe us an array of 3 values
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub colorRow(ByRef arrCatergory, strRange As String, Optional lReset)

    ' strRange A:B
    
    'Range("A" & CStr(intRowCount) & ":" & strCol & CStr(intRowCount)).Select
    If G_boolDevelopmentMode Then
        Range(strRange).Select
    End If
    
    If lReset Then
        arrCatergory(0, DBLFORECOLOR) = vbBlack
        arrCatergory(0, DBLBACKCOLOR) = xlNone
    End If
    
    With Range(strRange)
        ' back
        With .Interior
            .Pattern = xlSolid '
            .Color = arrCatergory(0, DBLBACKCOLOR)
        End With
    
        With .Font
            .Color = arrCatergory(0, DBLFORECOLOR)
            .TintAndShade = 0
        End With
    End With
   
End Sub

' --------------------------------------------------------------------------------------
' Procedure | createWithTimesheet
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub createWithTimesheet()
    Call destroyToolBar
    Call createToolBar
End Sub

' --------------------------------------------------------------------------------------
' Procedure | dayCheck
' Author    | Grant Bishop
' Date      | 23/03/2018
' Purpose   | Checks is day of week: Mon-Fri
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function dayCheck(strValue As String)
Dim strDayOfWeek As String
    If Not IsEmpty(strValue) Then
        If Len(Trim(strValue)) = 10 Then
            If IsDate(strValue) Then
                strDayOfWeek = Format(CDate(strValue), "dddd")
                dayCheck = (strDayOfWeek = "Monday" Or strDayOfWeek = "Tuesday" Or strDayOfWeek = "Wednesday" Or strDayOfWeek = "Thursday" Or strDayOfWeek = "Friday")
            End If
        End If
    End If
End Function

' --------------------------------------------------------------------------------------
' Procedure | decideColour
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function decideColour(strValue As String, ByRef arrCatergory) As Boolean

    decideColour = False
    
    strValue = UCase(strValue)
    
    If InStr(strValue, UCase("QlikView")) > 0 Then
        arrCatergory(0, 0) = "QlikView"
        arrCatergory(0, DBLBACKCOLOR) = 5287936
        arrCatergory(0, DBLFORECOLOR) = vbBlack
    End If
        
    If InStr(strValue, UCase("7CRM")) Then
        arrCatergory(0, 0) = "7CRM"
        arrCatergory(0, DBLBACKCOLOR) = vbBlue
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    If InStr(strValue, UCase("General")) Or InStr(strValue, UCase("Meetings:")) Then
        arrCatergory(0, 0) = "General Admin"
        arrCatergory(0, DBLBACKCOLOR) = 5296274
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    If InStr(strValue, UCase("eCas")) Then
        arrCatergory(0, 0) = "Development"
        arrCatergory(0, DBLBACKCOLOR) = 15040246
        arrCatergory(0, DBLFORECOLOR) = vbBlack
    End If
    
    If InStr(strValue, UCase("Cognos")) Or InStr(strValue, UCase("UK Apps")) Or InStr(strValue, UCase("Maximo")) Then
        arrCatergory(0, 0) = "Cognos"
        'arrCatergory(0, DBLBACKCOLOR) = 15773696
        'arrCatergory(0, DBLBACKCOLOR) = vbCyan
        arrCatergory(0, DBLBACKCOLOR) = vbMagenta
        'arrCatergory(0, DBLBACKCOLOR) = vbGreen    ' vbCyan
        arrCatergory(0, DBLFORECOLOR) = vbBlack
    End If
    
    If InStr(strValue, UCase("Group Treasury")) Then
        arrCatergory(0, 0) = "Group Treasury"
        arrCatergory(0, DBLBACKCOLOR) = 12611584
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    If InStr(strValue, UCase("BMS")) Then
        arrCatergory(0, 0) = "BMS"
        arrCatergory(0, DBLBACKCOLOR) = 49407
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    If InStr(strValue, UCase("Holiday:")) Or InStr(strValue, UCase("In Work")) Or InStr(strValue, UCase("!Lunch")) Then
        arrCatergory(0, 0) = "Holiday"
        arrCatergory(0, DBLBACKCOLOR) = vbBlack
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    'PAPPS / Support
    If InStr(strValue, UCase("Service Now")) Then
        arrCatergory(0, 0) = "System Admin"
        arrCatergory(0, DBLBACKCOLOR) = vbCyan
        arrCatergory(0, DBLFORECOLOR) = vbBlack
    End If

    If InStr(strValue, UCase("System")) Then
        arrCatergory(0, 0) = "System Admin"
        arrCatergory(0, DBLBACKCOLOR) = 10498160
        arrCatergory(0, DBLFORECOLOR) = vbWhite
    End If
    
    If InStr(strValue, UCase("Training")) Then
        arrCatergory(0, 0) = "Training"
        arrCatergory(0, DBLBACKCOLOR) = vbYellow
        arrCatergory(0, DBLFORECOLOR) = vbBlack
    End If
    
    If Trim(Len(arrCatergory(0, 0))) > 0 Then
        decideColour = True
    End If

End Function

' --------------------------------------------------------------------------------------
' Procedure | fixDateCol
' Author    | Grant Bishop
' Date      | 10/04/2018
' Purpose   | ensures the date always points to top row
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub fixDateCol(Wks As Worksheet, intRowCount As Integer, intStartOfWeekRow As Integer)
Dim strDateCol As String
Dim strRangeDateCol As String

    ' =N3008

    'MsgBox "stop", , "1GVB1 10/04/2018"
    
    strDateCol = getColLetter(C_COL_DATE)
    
    strRangeDateCol = "=" & strDateCol & intStartOfWeekRow
    
    If G_boolDevelopmentMode Then
        Wks.Cells(intRowCount, C_COL_DATE).Activate
    End If
    
    Wks.Cells(intRowCount, C_COL_DATE).Value = strRangeDateCol

End Sub

' --------------------------------------------------------------------------------------
' Procedure | formatTimeRecording
' Author    | Grant Bishop
' Date      | 10/04/2018
' Purpose   | will format the entire dynamic range
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub formatTimeRecording(Wks As Worksheet, intRowStart As Integer, intRowEnd As Integer, strStyle As String)
Dim strAddress As String

    ' A2793:M2810
    strAddress = "A" & intRowStart & ":M" & intRowEnd
    
    
    With Wks.Range(strAddress)
    
        'If G_boolDevelopmentMode Then
        '    .Activate
        'End If
    
        Select Case strStyle
            Case "CentreBoth"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            
            With .Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
    
            
        Case Else

            On Error Resume Next
            .Style = strStyle
            If Err.Number > 0 Then MsgBox " Style not defined! ::= " & strStyle
            On Error GoTo 0

        End Select
    
    End With
    
End Sub

' --------------------------------------------------------------------------------------
' Procedure | repairTimeRecording
' Author    | Grant Bishop
' Date      | 10/04/2018
' Purpose   | Loop along columns and fix up
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub repairTimeRecording(Wks As Worksheet, intRowStart As Integer, intRowEnd As Integer)
Dim intDeltaA As Integer
Dim intCol_S As Integer
Dim intCol_E As Integer
Dim strDeltaCol As String
Dim strDeltaRange As String
Dim strSumString As String

    intCol_S = getColNumber("A")
    intCol_E = getColNumber("M")
    
    strSumString = "=SUM(A" & CStr(intRowStart) & ":A" & CStr(intRowEnd) & ")"
    With Wks
    
        For intDeltaA = intCol_S To intCol_E
        
            strDeltaCol = getColLetter(intDeltaA)
            
            strDeltaRange = strDeltaCol & CStr(intRowStart) & ":" & strDeltaCol & CStr(intRowEnd)
            strSumString = "=SUM(" & strDeltaRange & ")"
            
            If G_boolDevelopmentMode Then
                Debug.Print strSumString
                .Cells(intRowEnd + 1, intDeltaA).Activate
            End If
            .Cells(intRowEnd + 1, intDeltaA).Value = strSumString
        
        Next intDeltaA
        
        ' finally do S
        strDeltaCol = "S"
        intDeltaA = getColNumber(strDeltaCol)
        
        strDeltaRange = strDeltaCol & CStr(intRowStart) & ":" & strDeltaCol & CStr(intRowEnd)
        strSumString = "=SUM(" & strDeltaRange & ")"
        
        If G_boolDevelopmentMode Then
            Debug.Print strSumString
            .Cells(intRowEnd + 1, intDeltaA).Activate
        End If
        .Cells(intRowEnd + 1, intDeltaA).Value = strSumString
        
    
    End With
    

End Sub

' --------------------------------------------------------------------------------------
' Procedure | returnLastCol
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     | move to common later
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function returnLastCol() As Integer

    returnLastCol = Cells(1, 1).End(xlToRight).Column

End Function

' --------------------------------------------------------------------------------------
' Procedure | returnLastColOnLastRow
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     | move to common later
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function returnLastColOnLastRow() As Integer

    'returnLastCol = Cells(1, 1).End(xlToRight).Column
    returnLastColOnLastRow = Cells(returnLastRow, 1).End(xlToRight).Column

End Function

' --------------------------------------------------------------------------------------
' Procedure | returnLastRow
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   | move to common later, code for PMSR - auto populating specifc areas,
'           | which can alos be used on my timesheet  dontcha'know
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function returnLastRow() As Integer
    returnLastRow = Cells(1, 1).SpecialCells(xlLastCell).Row

End Function

' --------------------------------------------------------------------------------------
' Procedure | sortHours
' Author    | Grant Bishop
' Date      | 23/03/2018
' Purpose   |
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub sortHours(Wks As Worksheet, intRowCount As Integer, intStartOfWeekRow As Integer)
Dim strSearchCat As String
Dim intTargetCol As Integer
Dim strTotalRange As String
Dim strTotalRangeSum As String
Dim strRange As String


    With Wks
        .Cells(intRowCount, C_COL_CATEGORY).Activate
        strSearchCat = Trim(.Cells(intRowCount, C_COL_CATEGORY).Value)
        
        If strSearchCat = "!NON WORKING" Then
            .Cells(intRowCount, C_COL_TOTAL).Value = ""
            
        Else
        
            ' need to put hours sum in if not there
            
            If Len(Trim(.Cells(intRowCount, C_COL_TOTAL).Value)) = 0 Then
            
                strTotalRange = "P" & intRowCount & "-O" & intRowCount
            
                strTotalRangeSum = "=SUM(" & strTotalRange & ")"
                .Cells(intRowCount, C_COL_TOTAL).Value = strTotalRangeSum
                
                ' =SUM(P3016-O3016)
                
            End If
            
            strRange = "A" & CStr(intRowCount) & ":" & "M" & CStr(intRowCount)
            '.Range(strRange).Activate
            .Range(strRange).ClearContents
        
            strSearchCat = transformCat(strSearchCat)
        
            If Len(Trim(strSearchCat)) > 0 Then
                
                intTargetCol = searchForValueInHeaderCol(Wks, strSearchCat, intStartOfWeekRow, "Whole", False)
                
                If intTargetCol > 0 Then
                    'Debug.Print intTargetCol
                    ' clear entire row
                    'strRange = "A" & CStr(intRowCount) & ":" & "M" & CStr(intRowCount)
                    '.Range(strRange).ClearContents
                    
                    .Cells(intRowCount, intTargetCol).Value = "=S" & CStr(intRowCount)
                
                Else
                    MsgBox "undefined - check", , "1GVB1 26/03/2018"
                
                End If
            Else
                ' clean out times
                strRange = "A" & CStr(intRowCount) & ":" & "M" & CStr(intRowCount)
                .Range(strRange).ClearContents
            
            End If
        End If
        
    End With

End Sub

' --------------------------------------------------------------------------------------
' Procedure | startOfWeekCheck
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   | For setting the range: InternalParameters.rangeTimeSheetRowNo()
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub startOfWeekCheck(Wks As Worksheet, intRowCount As Integer)
    With Wks
        If .Cells(intRowCount, 14).Value = "Date" Then
            If .Cells(intRowCount, 2).Value = "Week" Then
                If intRowCount <> getRangeValue("rangeTimeSheetRowNo") Then
                    Call setRangeValue("rangeTimeSheetRowNo", intRowCount)
                End If
            End If
        End If
    End With
End Sub

' --------------------------------------------------------------------------------------
' Procedure | TimeSheetSnapIn_ClearContents
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   |
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub TimeSheetSnapIn_ClearContents()
Dim dtToday As Date
Dim intCol As Currency
Dim intStartRow As Integer

    dtToday = Date
    
    intCol = getColNumber("N")
    
    intStartRow = searchForValueCols(ActiveSheet, CStr(dtToday), intCol)
    
    If intStartRow > 0 Then
        MsgBox "ready to start", , "1gvb12"
    
    Else
        MsgBox "unable to find row ", , "1gvb12"
    End If
    

    
End Sub

' --------------------------------------------------------------------------------------
' Procedure | timeSheetSnapIn_ColorCode
' Author    | Grant Bishop
' Date      | 08/05/2018
' Purpose   | Going to scan down descripiton - colouring in
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub timeSheetSnapIn_ColorCode()
Dim WksActiveSheet As Worksheet
Dim intLastRow As Currency
Dim intRowCount As Currency
Dim intCol As Integer
Dim strRange As String
'Dim intCount As Currency
Dim strExamin As String
Dim arrCatergory(0, 2)
Dim lReset As Boolean

    lReset = getRangeValue("rangeReset")
    ' reset flag to black and white
    
    ' Application.ScreenUpdating = False
    ' Application.ScreenUpdating = True

    ActiveCell.SpecialCells(xlLastCell).Select
    intLastRow = ActiveCell.Row
    
    strRange = ""
    intRowCount = 2                                                                     ' From the top
    intCol = 1                                                                         ' Q
    'intCount = 1
    G_boolDevelopmentMode = False
    
    ' intLastRow = 500
    
    Application.ScreenUpdating = G_boolDevelopmentMode
    
    Set WksActiveSheet = ActiveWorkbook.ActiveSheet                                     ' instantiate the sheet
    
    With WksActiveSheet
    
        'Do While intRowCount <= intLastRow
        Do While Len(Trim(.Range("A" & intRowCount).Value)) > 0
        
            If G_boolDevelopmentMode Then
                .Cells(intRowCount, intCol).Activate
            End If
            
            strExamin = .Cells(intRowCount, intCol).Value
            
            strRange = "A" & CStr(intRowCount) & ":B" & CStr(intRowCount)
            
            If decideColour(strExamin, arrCatergory) Then
                Call colorRow(arrCatergory, strRange, lReset)
            End If
            
            arrCatergory(0, 0) = ""
            intRowCount = intRowCount + 1
            
        Loop
        
    End With
    Set WksActiveSheet = Nothing
End Sub

' --------------------------------------------------------------------------------------
' Procedure | timeSheetSnapIn_writeTimeSheet
' Author    | Grant Bishop
' Date      | 23/03/2018
' Purpose   |
' To Do     | going to add in timings
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub timeSheetSnapIn_writeTimeSheet()
Dim Wks As Worksheet
Dim intLastRow As Integer
Dim intLastCol As Integer
Dim intRowCount As Integer
Dim intScanCol1 As Integer
Dim intExaminCol As Integer
Dim strDay As String
Dim intStartOfWeekRow As Integer
Dim intDeletaCount As Integer
Dim intDeltaFlag As Integer
Dim boolGetRowNo As Boolean

    intDeltaFlag = 1
    intDeletaCount = 1

    ' find last row to start on
    Set Wks = ActiveSheet
    
    intLastRow = getLastRow(Wks)
    intLastCol = getLastCol(Wks)
    intExaminCol = getColNumber("N")
    
    intRowCount = getRangeValue("rangeTimeSheetRowNo")
    boolGetRowNo = IIf(getRangeValue("rangeTimeSheetGetRowNo") = "Y", True, False)
    
    G_boolDevelopmentMode = False
    'G_boolDevelopmentMode = True
    ' intRowCount = 2142
    
    If intRowCount = 0 Then intRowCount = 1

    intScanCol1 = 1
    
    Application.ScreenUpdating = G_boolDevelopmentMode
        
    With Wks
    
        If G_boolDevelopmentMode Then
            .Activate
        End If
        
        Do While intRowCount <= intLastRow
            If G_boolDevelopmentMode Then
                .Cells(intRowCount, intExaminCol).Activate
            End If
            
            
            If boolGetRowNo Then
                Call startOfWeekCheck(Wks, intRowCount)
            End If
            
            
            If dayCheck(.Cells(intRowCount, intExaminCol).Value) Then           ' is it a valid day
            
                intStartOfWeekRow = intRowCount                                 ' will need this to build range +1
                
                strDay = .Cells(intRowCount, intExaminCol).Value
                intRowCount = intRowCount + 1                                   ' jump over the header
                
                While .Cells(intRowCount, intExaminCol).Value = strDay
                    
                    Call sortHours(Wks, intRowCount, intStartOfWeekRow)
                    Call fixDateCol(Wks, intRowCount, intStartOfWeekRow)
        
                    intRowCount = intRowCount + 1
        
                Wend
                
                intRowCount = intRowCount - 1
                
                Call repairTimeRecording(Wks, intStartOfWeekRow + 1, intRowCount)
                
                ' use a caller function to format as appropiate
                Call formatTimeRecording(Wks, intStartOfWeekRow + 1, intRowCount, "CentreBoth")
                
                'MsgBox "addValidationToColumn", , "1GVB1 11/04/2018"
                
                Call addValidationToColumn(Wks, "Q", intStartOfWeekRow + 1, intRowCount, "=rangeCategory")
                
                
                intDeltaFlag = intDeltaFlag + 1
                    
            End If
            
            intRowCount = intRowCount + 1
            
            'If intDeltaFlag = 6 Then
            '    intDeletaCount = intDeletaCount + 1
            '    If intDeletaCount = 9 Then
            '        intDeletaCount = 1
            '        intDeltaFlag = 1
            '    End If
            'End If
            
            'If intRowCount = 3116 Then
            '    G_boolDevelopmentMode = True
            '    Application.ScreenUpdating = True
            '    MsgBox "", , "1GVB1 04/04/2018"
            'End If
        Loop
    
    End With

End Sub

' --------------------------------------------------------------------------------------
' Procedure | transformCat
' Author    | Grant Bishop
' Date      | 23/03/2018
' Purpose   | coverts vertical descriptons to horizontal title headers
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Function transformCat(strSearchCat) As String
    ' Exceptions
    If strSearchCat <> "!NON WORKING" Then
        If strSearchCat = "Treasury" Then
            strSearchCat = "GT"
        ElseIf strSearchCat = "SAP" Then
            strSearchCat = "UK Apps"
        ElseIf strSearchCat = "ProArc" Then
            strSearchCat = "UK Apps"
        ElseIf strSearchCat = "Maximo" Then
            strSearchCat = "UK Apps"
        'ElseIf strSearchCat = "System" Then
        '    strSearchCat = "System Admin"
        ElseIf strSearchCat = "Service Now" Then
            strSearchCat = "System"
        ElseIf strSearchCat = "Trello" Then
            strSearchCat = "General"
        ElseIf strSearchCat = "Personal" Or strSearchCat = "!Holiday" Then
            strSearchCat = "OOO"
        End If
        
        transformCat = strSearchCat
        
    End If
    
End Function

