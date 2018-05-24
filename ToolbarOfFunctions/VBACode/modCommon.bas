Attribute VB_Name = "modCommon"
Option Explicit

' Objects
Dim objFileSystem

Global G_boolDevelopmentMode As Boolean             ' used for development mode: toggle between true and false to turn on watch points
Global Const GC_MSGBOX = "Toolbar of Functions"

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | allTrim
' Author    | G Bishop
' Date      | 14/09/2016
' Purpose   | git rid of special characters
' Sub Test_AllTrim()
'     Debug.Print allTrim(Chr(160) & Chr(10))
'     Debug.Print allTrim(Chr(10) & Chr(160))
' End Sub
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 14/09/2016 | Added in this header
' G Bishop  | 04-04-2018 | Didnt take into account spaces - beleive it or not
' ----------+------------+----------------------------------------------------------------------------------------
Function allTrim(ByVal strValue) As String
    If Len(Trim(strValue)) > 0 Then
        strValue = Trim(stripOutChar(strValue, Chr(10)))
        If Len(Trim(strValue)) > 0 Then
            strValue = Trim(stripOutChar(strValue, Chr(160)))
        End If
    Else
        strValue = Trim(strValue)
    End If
    allTrim = strValue
End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub autoFit(WksActiveSheet As Worksheet, strDoWhat As String)
    With WksActiveSheet
        If strDoWhat = "USERS" Then
            .Range("A1:D1").Font.Bold = True
            .Columns("A:A").EntireColumn.autoFit
            .Columns("B:B").EntireColumn.autoFit
            .Columns("C:C").EntireColumn.autoFit
            .Columns("D:D").EntireColumn.autoFit
        ElseIf strDoWhat = "GROUPS" Then
            .Range("A1:D1").Font.Bold = True
            .Columns("A:A").EntireColumn.autoFit
            .Columns("B:B").EntireColumn.autoFit
            .Columns("C:C").EntireColumn.autoFit
            .Columns("D:D").EntireColumn.autoFit
        End If
    End With
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub clearCells(WksActiveSheet As Worksheet, intNoOfCells)
Dim intCount As Currency
    For intCount = 1 To intNoOfCells
        With WksActiveSheet.Cells(1, intCount)
            .Select
            .Value = ""
            .EntireColumn.EntireColumn.autoFit
            .Font.Bold = False
        End With
    Next intCount
End Sub

'--------------------------------------------------------------------------
' Clear any Formatting
'--------------------------------------------------------------------------
Sub clearFormatting()

    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Range("A1").Select
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | countCharacters
' Author    | G Bishop
' Date      | 30/01/2017
' Purpose   | Counts the number of occurrences of: strChar in: strValue - with a loop
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function countNoOfCharacters1(strValue, strChar)
Dim intCount
Dim intNo
    intNo = 0
    For intCount = 1 To Len(strValue)
        If Mid(strValue, intCount, 1) = strChar Then
            intNo = intNo + 1
        End If
    Next
    countNoOfCharacters1 = intNo

End Function

' Extract file name from the path
Function extractFileNameOnly(strFileName As String) As String
    While InStr(strFileName, "\") > 0
        strFileName = Mid(strFileName, (InStr(1, strFileName, "\", vbTextCompare) + 1))
    Wend
    extractFileNameOnly = strFileName
End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | fileExists( strFileName )
' Author    | G Bishop
' Date      | 30/01/2017
' Purpose   | Uses file system object to check if a file exists
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function fileExists(strFileName)
'Dim objFileSystem
    If (Not IsNull(objFileSystem)) Then Set objFileSystem = CreateObject("Scripting.FileSystemObject")               ' instantiate the file system object
    fileExists = objFileSystem.fileExists(strFileName)
    'Set objFileSystem = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : formatCells
' Author    : Grant Bishop
' Date      : 06/10/2015
' Purpose   : will format the cell to a set of specific rules
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Sub formatCells(Wks As Worksheet, intRow As Integer, intCol As Integer, strStyle As String)

    With Wks.Cells(intRow, intCol)
    
        If strStyle = "Normal" Then
            .Style = "Normal"
            '.HorizontalAlignment = xlLeft
            .HorizontalAlignment = xlCenter
            
        ElseIf strStyle = "Small" Then
            .Style = "Normal"
            .Font.Size = 8
            
        ElseIf strStyle = "SmallCentred" Then
            .Style = "Normal"
            .Font.Size = 8
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        Else
            .Style = strStyle
            .Font.Size = 8
            
        End If
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------
'Function: getColLetter
'
'Description:
'   returns 'A' for col 1, 'B' for col 2 ,..
'Version Customisations:
'
'To Do:
'
'Modification History:
'----------------------------------------------------------------------------------------------------------------
Function getColLetter(intCol As Integer) As String
Dim strCol As String

    strCol = Cells(1, intCol).Address
    If (Len(strCol) > 4) Then
        strCol = Mid(strCol, Len(strCol) - 3, 2)
    Else
        strCol = Mid(strCol, Len(strCol) - 2, 1)
    End If

    getColLetter = strCol

End Function

'----------------------------------------------------------------------------------------------------------------
'Function: getColNumber
'
'Description:
'   convert's column string to column numbers: A->1, B->2, ...
'Version Customisations:
'
'To Do:
'
'Modification History:
'----------------------------------------------------------------------------------------------------------------
Public Function getColNumber(strCol As String) As Integer
Dim intOutNum As Integer
Dim intLen As Integer
Dim intI As Integer

    intLen = Len(strCol)

    For intI = 1 To intLen
        intOutNum = (Asc(UCase(Mid(strCol, intI, 1))) - 64) + intOutNum * 26
    Next intI

    getColNumber = intOutNum

End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Function getDN(strName As String, strDomain As String)
'Dim wshShell, wshNetwork
Dim objTrans, objDomain
Dim strRetValue
' Constants for the NameTranslate object.
Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1

    Set objTrans = CreateObject("NameTranslate")
    Set objDomain = GetObject("LDAP://rootDse")
    objTrans.Init ADS_NAME_INITTYPE_GC, ""
    'objTrans.Set ADS_NAME_TYPE_NT4, wshNetwork.UserDomain & "\" & strName & "$"
    'objTrans.Set ADS_NAME_TYPE_NT4, "GROUP\" & strName
    
    objTrans.Set ADS_NAME_TYPE_NT4, strDomain & "\" & strName
    strRetValue = objTrans.GET(ADS_NAME_TYPE_1779)
        
    getDN = strRetValue
    
    Set objTrans = Nothing
    Set objDomain = Nothing
    
End Function

Sub getextendedDetailsKeepToOneSide()
Dim arrHeaders(35)
Dim objFolder
Dim i
Dim strFileName
Dim objShell

    ' cana I get for just one file?
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace("A:\5.2.12\5.2.12\Africa\1 Received")


    For i = 0 To 34
    
        arrHeaders(i) = objFolder.getDetailsOf(objFolder.Items, i)
    Next
    
    For Each strFileName In objFolder.Items
    
        Debug.Print strFileName
        Debug.Print objFolder.getDetailsOf(strFileName, 0)
        
        If strFileName = objFolder.getDetailsOf(strFileName, 0) Then
            Debug.Print objFolder.getDetailsOf(strFileName, 22)
        End If
    
        For i = 0 To 34
            Debug.Print i & vbTab & arrHeaders(i) & ": " & objFolder.getDetailsOf(strFileName, i)
        Next
    Next

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | M_xCommon.getLastCol
' Author    | Grant Bishop
' Date      | 01/10/2015
' Purpose   | gets the last selectable column in the supplied worksheet
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Function getLastCol(Wks As Worksheet) As Integer
    getLastCol = Wks.Cells(1, 1).SpecialCells(xlLastCell).Column

End Function

'---------------------------------------------------------------------------------------
' Procedure : getLastCol
' Author    : Grant Bishop
' Date      : 23/03/2018
' Purpose   : unsure what this is doing
'---------------------------------------------------------------------------------------
Function getLastCol_Odd() As Integer

    getLastCol_Odd = Cells(1, 1).End(xlToRight).Column

End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | getLastRow
' Author    | G Bishop
' Date      | 01/10/2015
' Purpose   | uses excel calls to get the last row - there are a choice of 4
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Function getLastRow(Wks As Worksheet, Optional strMethod As String, Optional strCell As String) As Long
Dim lngLastRow As Long

    If strMethod = "" Then strMethod = "A"

    If strMethod = "A" Then
        lngLastRow = Wks.Cells.SpecialCells(xlLastCell).Row
    ElseIf strMethod = "B" Then
        If strCell = "" Then strCell = "A1"
        
        With Wks
            lngLastRow = .Range(strCell).SpecialCells(xlCellTypeLastCell).Row
        End With
        
    ElseIf strMethod = "C" Then
        With Wks.UsedRange
            lngLastRow = .Rows(.Rows.Count).Row
        End With

    End If

    getLastRow = lngLastRow

End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | getPathFileDriveOFS
' Author    | G Bishop
' Date      | 01/10/2015
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Function getPathFileDriveOFS(strFileName, strReturnWhat)
Dim strRetValue
Dim objFile

    If (Not IsNull(objFileSystem)) Then Set objFileSystem = CreateObject("Scripting.FileSystemObject")

    If LCase(Left(strFileName, 5)) = LCase("http:") Then
        strFileName = Replace(Mid(strFileName, 6, Len(strFileName)), "/", "\")
    End If

    If fileExists(strFileName) Then

        Set objFile = objFileSystem.GetFile(strFileName)

        If strReturnWhat = "Drive" Then                                     ' return drive letter
            strRetValue = Mid(objFileSystem.GetParentFolderName(objFile), 1, 2)
        ElseIf strReturnWhat = "Path" Then                                  ' return path part
            strRetValue = objFileSystem.GetParentFolderName(objFile)
        ElseIf strReturnWhat = "File" Then                                  ' return full filename
            strRetValue = objFileSystem.GetFileName(objFile)
        ElseIf strReturnWhat = "FilePart" Then                              ' return filename w/out extension
            strRetValue = objFileSystem.GetBaseName(objFile)
        ElseIf strReturnWhat = "Ext" Then                                   ' return extension
            strRetValue = objFileSystem.GetExtensionName(objFile)
        End If
    Else                                                                    ' file doesn't actually exist - so return what we can from string supplied

        If InStr(strFileName, "\") > 0 Then                                 ' file has a URL
            If strReturnWhat = "Drive" Then                                 ' return drive letter
                strRetValue = Mid(strFileName, 1, 2)
            ElseIf strReturnWhat = "Path" Then                              ' return path part
                strRetValue = Mid(strFileName, 1, (Len(strFileName) - Len(getValueFromRight(strFileName, "\"))) - 1)
            End If
        End If

        If strReturnWhat = "File" Then                                      ' return full filename
            strRetValue = getValueFromRight(strFileName, "\")
            If Len(Trim(strRetValue)) = 0 Then strRetValue = strFileName    ' if no value return itself

        ElseIf strReturnWhat = "FilePart" Then                              ' return filename w/out extension
            If InStr(strFileName, "\") = 0 Then                             ' does it have a path?
                strRetValue = Left(strFileName, Len(strFileName) - CInt(Len(getValueFromRight(strFileName, ".")) + 1))
            Else                                                            ' a reclusive call to get file from file\path
                strRetValue = Left(getPathFileDriveOFS(strFileName, "File"), Len(getPathFileDriveOFS(strFileName, "File")) - CInt(Len(getValueFromRight(getPathFileDriveOFS(strFileName, "File"), ".")) + 1))
            End If

        ElseIf strReturnWhat = "Ext" Then                                   ' return extension
            strRetValue = getValueFromRight(strFileName, ".")

        End If

    End If

    getPathFileDriveOFS = strRetValue

End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | getRangeValue
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function getRangeValue(strWhatRange) As String
    getRangeValue = InternalParameters.Range(strWhatRange).Value
End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | getValue( strFileName, strWhat )
' Author    | G Bishop
' Date      | 27/02/2017
' Purpose   | walk back till strWhat
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function getValueFromRight(strFileName, strWhat)
Dim strExtract
Dim strRetValue
Dim intCount
    If countNoOfCharacters1(strFileName, strWhat) > 0 Then
        For intCount = Len(strFileName) To 1 Step -1
            strExtract = Mid(strFileName, intCount, 1)
            If strExtract = strWhat Then
                Exit For
            End If

            strRetValue = strExtract & strRetValue
        Next
    End If
    getValueFromRight = strRetValue
End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | handleRangesSub
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Sub handleRangesSub(strDoWhat As String, strWhatRange As String, ByRef strValue)
    With InternalParameters
        Select Case strDoWhat
        Case "Set"
            InternalParameters.Range(strWhatRange).Value = strValue
        Case "Get"
            strValue = InternalParameters.Range(strWhatRange).Value
        Case "Forget"
        End Select
    End With
End Sub

'--------------------------------------------------------------------------
' Check for the end of the sheet - unsure if this is still relevant
'--------------------------------------------------------------------------
Function isSheetEnd(lngX) As Boolean
    If (lngX + 1) < 65537 Then
        isSheetEnd = False
    Else
        isSheetEnd = True
        MsgBox "this routine can only handle up to 65537 rows - rewrite required?"
    End If
End Function

'------------------------------------------------------------------------------------------
' v = fso.DeleteFile(strFileSpec, [bForce])   'Force = True will delete ReadOnly.
'------------------------------------------------------------------------------------------
Function killFile(strFileName)
    'Err.Clear: On Error Resume Next
    killFile = objFileSystem.DeleteFile(strFileName, True)
    'If Err.Number <> 0 Then Call MyErrorHandler("KillFile", " Could not delete file: " & strFileName): On Error GoTo 0
End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Sub myDisplayMessage(strSayWhatMessage As String, strDoWhat As String)

    If strDoWhat = "On" Then
        Application.DisplayStatusBar = True
    ElseIf strDoWhat = "Display" Then
        Application.StatusBar = strSayWhatMessage
    ElseIf strDoWhat = "Off" Then
        Application.DisplayStatusBar = False
    End If
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub NavigateHome(WksActiveSheet As Worksheet)
Dim WksRememeberSheet As Worksheet
    Set WksRememeberSheet = ActiveSheet
    WksActiveSheet.Activate
    On Error Resume Next
    WksActiveSheet.Range("A1").Select
    On Error GoTo 0
    WksRememeberSheet.Activate
End Sub

'--------------------------------------------------------------------------
' Check were not on the parameters sheet
'--------------------------------------------------------------------------
Function notParameters(Wks As Worksheet) As Boolean
    notParameters = (Wks.Name <> "InternalParameters")
End Function

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | populateCDCmboBox
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Sub populateCDCmboBox(ByRef cmboDifferences)
'Public Sub populateCDCmboBox(strSheetName As String, ByRef cmboDifferences)
'Dim WksParams As Worksheet
Dim strValue As String

    With cmboDifferences
        .Clear
        .AddItem "Colour"
        .AddItem "Clear"
    End With
    
    Call handleRangesSub("Get", "rangeCompareOption", strValue)
    
    cmboDifferences.Value = strValue

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | populateHDCmboBox
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Sub populateHDCmboBox(ByRef cmboHighLightOrDelete)
'Public Sub populateHDCmboBox(strSheetName As String, ByRef cmboHighLightOrDelete)
Dim strValue As String

    With cmboHighLightOrDelete
        .Clear
        .AddItem "Delete"
        .AddItem "Highlight"
        .AddItem "ClearCell"
    End With
    
    Call handleRangesSub("Get", "rangeHighlightOrDeleteOption", strValue)
    
    cmboHighLightOrDelete.Value = strValue

End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Sub populateModeCmboBox(ByRef cmboModeAorB)
'Public Sub populateModeCmboBox(strSheetName As String, ByRef cmboModeAorB)
'Dim WksParams As Worksheet
Dim strValue As String

    With cmboModeAorB
        .Clear
        .AddItem "A"
        .AddItem "B"
    End With
    
    Call handleRangesSub("Get", "rangeDelBlankLinesModeAorB", strValue)
    
    cmboModeAorB.Value = strValue

End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Sub populateTBCmboBox(ByRef cmboShowToolbarDescription)
'Public Sub populateTBCmboBox(strSheetName As String, ByRef cmboShowToolbarDescription)
'Dim WksParams As Worksheet
Dim strValue As String

    'Set WksParams = Sheets(strSheetName)
    
    With cmboShowToolbarDescription
        .Clear
        .AddItem "True"
        .AddItem "False"
    End With
    
    Call handleRangesSub("Get", "rangeShowDescriptionOption", strValue)
    
    cmboShowToolbarDescription.Value = strValue

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | populateTTCmboBox
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Public Sub populateTTCmboBox(ByRef cmboDisplayTimeTaken)
'Public Sub populateTTCmboBox(strSheetName As String, ByRef cmboDisplayTimeTaken)
'Dim WksParams As Worksheet
Dim strValue As String

    With cmboDisplayTimeTaken
        .Clear
        .AddItem "True"
        .AddItem "False"
    End With
    
    Call handleRangesSub("Get", "rangeTimeTaken", strValue)
    
    cmboDisplayTimeTaken.Value = strValue

End Sub

'--------------------------------------------------------------------------
' Check were not on the parameters sheet
'--------------------------------------------------------------------------
Function retrieveValue(strS, strSW, intNoS)
Dim intA As Integer
    intA = 1
    While intA < intNoS + 1
        strS = Mid(strS, InStr(strS, "," & strSW) + 1, Len(strS) - InStr(strS, "," & strSW))
        intA = intA + 1
    Wend
    retrieveValue = Right(Mid(strS, 1, InStr(strS, ",") - 1), Len(strSW) + 1)
End Function

'--------------------------------------------------------------------------
' return the row num
'--------------------------------------------------------------------------
Public Function searchForValue(WksSearchSheet As Worksheet, strSearchValue As String, intWhichScanCol As Currency) As Currency
Dim colSearchCol
Dim intRowNo As Currency
   
    With WksSearchSheet.Columns(intWhichScanCol)
        Set colSearchCol = .Find(Trim(strSearchValue), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not colSearchCol Is Nothing Then
            intRowNo = colSearchCol.Row
        End If
    End With
    
    searchForValue = intRowNo

End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : searchForValueCols
' Author    : Grant Bishop
' Date      : 08/02/2016
' Returns   : Currency cur
' Purpose   :
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Public Function searchForValueCols(WksSearchSheet As Worksheet, strSearchValue As String, intWhichScanCol As Currency) As Currency
'Dim colSearchCol
Dim intRowNo As Currency

    WksSearchSheet.Activate
    
    MsgBox "", , "1GVB12"
    MsgBox "", , "1GVB12"
    MsgBox "", , "1GVB12"
    MsgBox "", , "1GVB12"

    Columns("N:N").Select
    Columns(intWhichScanCol).Select
    
    Range("N2037").Activate
    Selection.Find(What:=strSearchValue, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate


    intRowNo = ActiveCell.Row
    
    searchForValueCols = intRowNo
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : searchForValueInHeaderCol
' Author    : Grant Bishop
' Date      : 23/03/2018
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function searchForValueInHeaderCol(Wks As Worksheet, strSearchValue As String, intWhichScanRow As Integer, strWholeOrPart As String, Optional boolSwitch As Boolean) As Integer
Dim intColNo As Currency
'Dim intLastRow As Integer
Dim boolFound As Boolean

    With Wks
    
        For intColNo = 1 To 14
            If Not IsError(.Cells(intWhichScanRow, intColNo).Value) Then
                If .Cells(intWhichScanRow, intColNo).Value <> "" Then
                    If strWholeOrPart = "Whole" Then
                        If UCase(.Cells(intWhichScanRow, intColNo).Value) = UCase(strSearchValue) Then
                            boolFound = True
                        End If
                    Else
                        If Not boolSwitch Then
                            If InStr(UCase(strSearchValue), UCase(.Cells(intWhichScanRow, intColNo).Value)) > 0 Then
                                boolFound = True
                            End If
                        Else
                            If InStr(UCase(.Cells(intWhichScanRow, intColNo).Value), UCase(strSearchValue)) > 0 Then
                                boolFound = True
                            End If

                        End If
                    End If

                End If
                If boolFound Then Exit For
            End If
        Next intColNo
    End With

    If boolFound Then
        searchForValueInHeaderCol = intColNo
    End If

End Function

Sub setFontBlank()
    Cells.Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub setRangeBold(strRangeAddr As String)
    With Range(strRangeAddr)
        .Select
        .Font.Bold = True
        .EntireColumn.EntireColumn.autoFit
    End With
End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | setRangeValue
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   |
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function setRangeValue(strWhatRange, varWhat) As String
    InternalParameters.Range(strWhatRange).Value = varWhat
End Function

'--------------------------------------------------------------------------
' Sort the sheet after populating
'--------------------------------------------------------------------------
Sub sortSheet(WksActiveSheet As Worksheet)
Dim intLastRow As Integer
Dim intLastCol As Integer

    With WksActiveSheet
    
        intLastRow = .Cells(1, 1).SpecialCells(xlLastCell).Row
        intLastCol = .Cells(1, 1).End(xlToRight).Column
        
        ' Selection.End(xlToRight).Select
        ' Range("A1:B" & CStr(intLastRow)).Select
        ' Range(1:intLastCol, intLastCol).Select
        
        
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A2:A" & CStr(intLastRow)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            
            .SetRange Range("A1:B" & CStr(intLastRow))
            
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End With
End Sub

'----------------------------------------------------------------------------------------------------------------
' Procedure : stripOutChar
' Author    : Grant Bishop
' Date      : 02/11/2015
' Returns   : String str
' Purpose   : removed passed values returns from cells being imported
'----------------------------------------------------------------------------------------------------------------
Function stripOutChar(ByVal strMyRange, ByVal strWhatChar) As String
Dim intChkCR As Integer
    While InStrRev(strMyRange, strWhatChar) > 0
        intChkCR = InStrRev(strMyRange, strWhatChar)
        strMyRange = Mid(strMyRange, intChkCR + 1, Len(strMyRange))
    Wend
    stripOutChar = strMyRange
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : stripOutChar
' Author    : Grant Bishop
' Date      : 02/11/2015
' Returns   : String str
' Purpose   : removed hard returns from cells being imported
'----------------------------------------------------------------------------------------------------------------
Function stripOutChar2(strMyRange As String) As String
Dim intChkCR As Integer
    While InStrRev(strMyRange, Chr(10)) > 0                 ' if the range contains any carriage returns, remove them
        intChkCR = InStrRev(strMyRange, Chr(10))
        strMyRange = Left(strMyRange, intChkCR - 1)
    Wend
    stripOutChar2 = strMyRange
End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub writeHeaders(WksActiveSheet As Worksheet, strDoWhat As String)
    
    Call clearCells(WksActiveSheet, 5)
    
    If strDoWhat = "FILES" Then
        With WksActiveSheet
            .Cells(1, 1).Value = "File Name"
            .Cells(1, 2).Value = "Date"
            .Cells(1, 3).Value = "Size"
            .Cells(1, 4).Value = "Version"
            .Cells(1, 5).Value = "File Name Extracted"
        End With
        Call setRangeBold("A1:E1")
    ElseIf strDoWhat = "USERS" Then
        With WksActiveSheet
            .Cells(1, 1).Value = "Group Name":
            .Cells(1, 2).Value = "Group Description"
        End With
        Call setRangeBold("A1:B1")
    ElseIf strDoWhat = "GROUP" Then
        With WksActiveSheet
            .Cells(1, 1).Value = "Name"
            .Cells(1, 2).Value = "Full Name"
            .Cells(1, 3).Value = "Description"
            .Cells(1, 4).Value = "AccountDisabled"
        End With
        Call setRangeBold("A1:D1")
    ElseIf strDoWhat = "ZAP" Then
        Columns("A:J").Select
        Selection.ColumnWidth = 8.38
    End If
    Call NavigateHome(ActiveSheet)
End Sub

