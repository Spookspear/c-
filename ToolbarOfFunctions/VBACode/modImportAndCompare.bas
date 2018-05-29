Attribute VB_Name = "modImportAndCompare"
Option Explicit
Dim objFileSystem       ' Objects
Dim objShell

Dim gnCount

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Function browseFolder(myStartLocation, blnSimpleDialog)
' This function generates a Browse Folder dialog
' and returns the selected folder as a string.
'
' Arguments:
' myStartLocation   [string]  start folder for dialog, or "My Computer", or
'                             empty string to open in "Desktop\My Documents"
' blnSimpleDialog   [boolean] if False, an additional text field will be
'                             displayed where the folder can be selected
'
' Returns:          [string]  the fully qualified path to the selected folder
'
' Based on the Hey Scripting Guys article
' "How Can I Show Users a Dialog Box That Only Lets Them Select Folders?"
' http://www.microsoft.com/technet/scriptcenter/resources/qanda/jun05/hey0617.mspx
'
' Function written by Rob van der Woude
' http://www.robvanderwoude.com
    Const MY_COMPUTER = &H11&
    Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt

    ' Set the options for the dialog window
    strPrompt = "Select a folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If

    ' Create a Windows Shell object
    Set objShell = CreateObject("Shell.Application")

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase(myStartLocation) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace(MY_COMPUTER)
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder(WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath)

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        browseFolder = ""
        Exit Function
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    browseFolder = objPath
End Function
by typing the fully qualified path
'
' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | compareSheets
' Author    | G Bishop
' Date      | 17/05/2017
' Purpose   | Altered to include number of columns to check
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop    17-05-2017   Lots
' G Bishop    01-05-2018   Will dynamically work out last col if X present
' ----------+------------+----------------------------------------------------------------------------------------
Public Sub compareSheets(Wks1 As Worksheet, Wks2 As Worksheet)
Dim intCount1 As Currency
Dim boolFound As Boolean
Dim strFile1 As String
Dim intScore As Integer
Dim curSheetLastRow1 As Currency
Dim curSheetLastRow2 As Currency
Dim curTargetRow As Currency
Dim curSourceRow As Currency
Dim intStartRow As Currency
Dim strDoWhat As String
Dim intNoCheckCols As Integer
Dim intDeltaCount As Integer
Dim curStartColumToCheck As Currency
Dim lngClrFoundFont, lngClrNotFoundFont As Long
Dim lngClrFoundBack, lngClrNotFoundBack As Long
Dim strNoCheckColsCheck As String

    G_boolDevelopmentMode = False

    ' Read in data from InternalParameters
    strDoWhat = getRangeValue("rangeCompareOption")

    strNoCheckColsCheck = getRangeValue("rangeNoOfColumnsToCheck")
    If strNoCheckColsCheck = "X" Then
        intNoCheckCols = getLastCol(Wks1)
    Else
        intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck")
    End If

    intStartRow = getRangeValue("rangeComparingStartRow")
    curStartColumToCheck = getRangeValue("rangeDupliateColumnToCheck")

    lngClrFoundFont = InternalParameters.Range("rangeColourFound").Font.Color
    lngClrNotFoundFont = InternalParameters.Range("rangeColourNotFound").Font.Color

    lngClrFoundBack = InternalParameters.Range("rangeColourFound").Interior.Color
    lngClrNotFoundBack = InternalParameters.Range("rangeColourNotFound").Interior.Color

    curSheetLastRow2 = getLastRow(Wks2)
    curSheetLastRow1 = getLastRow(Wks1)

    Call setFontBlank
    Wks1.Cells(intStartRow, curStartColumToCheck).Select

    intDeltaCount = 1

    ' start of loop
    For curSourceRow = intStartRow To curSheetLastRow1


        'If curSourceRow = 2234 Then
        '    MsgBox "stop", , "1GVB1 13/10/2017"
        'End If


        boolFound = False
        strFile1 = Wks1.Cells(curSourceRow, curStartColumToCheck).Value

        curTargetRow = searchForValue(Wks2, strFile1, curStartColumToCheck)

        If curTargetRow > 0 Then
            boolFound = True
            intScore = 0: intCount1 = 0

            ' start from correct column
            For intCount1 = curStartColumToCheck To (intNoCheckCols + curStartColumToCheck - 1)

                If G_boolDevelopmentMode Then
                    Wks1.Cells(curSourceRow, intCount1).Activate
                    Debug.Print LCase(Wks1.Cells(curSourceRow, intCount1).Value)
                    Debug.Print LCase(Wks2.Cells(curTargetRow, intCount1).Value)
                End If

                If LCase(Wks1.Cells(curSourceRow, intCount1).Value) = LCase(Wks2.Cells(curTargetRow, intCount1).Value) Then
                    intScore = intScore + 1
                End If
            Next intCount1

            ' Score system = if all the same then can blue it
            If intScore = intNoCheckCols Then
                For intCount1 = curStartColumToCheck To (intNoCheckCols + curStartColumToCheck - 1)
                    If strDoWhat = "Colour" Then
                        Wks1.Cells(curSourceRow, intCount1).Font.Color = lngClrFoundFont
                        'Wks1.Cells(curSourceRow, intCount1).Interior.Color = lngClrFoundBack

                    Else
                        Wks1.Cells(curSourceRow, intCount1).Value = ""
                    End If
                Next intCount1
            End If

        Else
            ' this wasn't found
            For intCount1 = curStartColumToCheck To (intNoCheckCols + curStartColumToCheck - 1)
                Wks1.Cells(curSourceRow, intCount1).Font.Color = lngClrNotFoundFont
                'Wks1.Cells(curSourceRow, intCount1).Interior.Color = lngClrNotFoundBack

            Next intCount1
        End If

        ' Every zoom 10% screen load '
        If intDeltaCount >= 5000 Then

            'MsgBox "work on later - when checking more data", , "1GVB1 17/05/2017"

            Application.ScreenUpdating = True
            Debug.Print "Row No: " & CStr(curSourceRow)
            Wks1.Cells(curSourceRow, intCount1).Select
            ActiveWindow.LargeScroll Down:=1
            Application.ScreenUpdating = False
            intDeltaCount = 1
        End If

        intDeltaCount = intDeltaCount + 1

    Next curSourceRow

End Sub

'--------------------------------------------------------------------------
' will get the file name from the loop
'--------------------------------------------------------------------------
Public Sub extractFileNameLoop(strSheetName)
Dim WksActiveSheet As Worksheet
Dim intCount As Currency
Dim intStartRow As Currency
Dim strFileName As String
Dim intRows As Currency
Dim objFile1 As Object

    ActiveCell.SpecialCells(xlLastCell).Select
    intRows = ActiveCell.Row

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    intStartRow = 2

    Set WksActiveSheet = ActiveWorkbook.Sheets(strSheetName)

    For intCount = intStartRow To intRows

        strFileName = WksActiveSheet.Cells(intCount, 1)

        If strFileName > "" Then
            Set objFile1 = objFileSystem.GetFile(strFileName)
            With objFile1
                WksActiveSheet.Cells(intCount, 2) = .DateLastModified
                WksActiveSheet.Cells(intCount, 3) = .Size
            End With
            WksActiveSheet.Cells(intCount, 4) = objFileSystem.GetFileVersion(strFileName)

            WksActiveSheet.Cells(intCount, 5) = extractFileNameOnly(strFileName)

        End If
    Next intCount

    Set objFile1 = Nothing

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | getDetailsOf
' Author    | G Bishop
' Date      | 17/05/2017
' Purpose   | this can be extended to get any 'Extended propertie'
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function getDetailsOf(strFileName As String, intWhichAttrib As Integer) As String
Dim objFolder
Dim objFile
Dim strPath
Dim strFile

    strPath = getPathFileDriveOFS(strFileName, "Path")
    strFile = getPathFileDriveOFS(strFileName, "File")

    ' move to global?
    Set objFolder = objShell.Namespace(strPath)

    For Each objFile In objFolder.Items

        If strFile = objFolder.getDetailsOf(objFile, 0) Then
            getDetailsOf = objFolder.getDetailsOf(objFile, intWhichAttrib)
            Exit For
        End If

    Next

    Set objFolder = Nothing

End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Sub grabDetails(strSheetName)
Dim WksActiveSheet As Worksheet
Dim intCount As Currency
Dim intStartRow As Currency
Dim strFileName As String
Dim intRows As Currency
Dim objFile1 As Object

    ActiveCell.SpecialCells(xlLastCell).Select
    intRows = ActiveCell.Row

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("Shell.Application")

    intStartRow = 2

    Set WksActiveSheet = ActiveWorkbook.Sheets(strSheetName)

    For intCount = intStartRow To intRows

        strFileName = WksActiveSheet.Cells(intCount, 1)

        If strFileName > "" Then
            Set objFile1 = objFileSystem.GetFile(strFileName)
            With objFile1
                WksActiveSheet.Cells(intCount, 2) = .DateLastModified
                WksActiveSheet.Cells(intCount, 3) = .Size
            End With

            WksActiveSheet.Cells(intCount, 4) = objFileSystem.GetFileVersion(strFileName)
            If Len(Trim(WksActiveSheet.Cells(intCount, 4))) = 0 Then
                WksActiveSheet.Cells(intCount, 4) = getDetailsOf(strFileName, 22)
            End If

            WksActiveSheet.Cells(intCount, 5) = extractFileNameOnly(strFileName)

        End If
    Next intCount

    Set objFile1 = Nothing

End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Function IsValue(obj)
' Check whether the value has been returned.
Dim tmp
    On Error Resume Next
    tmp = " " & obj
    If Err <> 0 Then
        IsValue = False
    Else
        IsValue = True
    End If
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------
' recursive function
'--------------------------------------------------------------------------
Function listSubFoldersAndFiles(strSubFolderPath, strOutputSheet, intHighlightRows)
Dim objFolder
Dim ObjFileCollection
Dim ObjSubFolderCollection
Dim ObjSubFolder
Dim objFile

    Set objFolder = objFileSystem.GetFolder(strSubFolderPath)
    Set ObjFileCollection = objFolder.Files

    For Each objFile In ObjFileCollection
        ' Stick the value on the names sheet
        ActiveWorkbook.Sheets(strOutputSheet).Cells(gnCount, 1).Value = objFile.Path

        If Len(objFile.Path) > intHighlightRows And intHighlightRows <> 0 Then
            ActiveWorkbook.Sheets(strOutputSheet).Cells(gnCount, 1).Font.Color = vbRed
        End If

        gnCount = gnCount + 1
    Next

    Set ObjSubFolderCollection = objFolder.SubFolders
    For Each ObjSubFolder In ObjSubFolderCollection
        Call listSubFoldersAndFiles(ObjSubFolder.Path, strOutputSheet, intHighlightRows)
    Next

End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Function populateSheetFromFolder(strOutputSheet, intHighlightRows) As Boolean
Dim strInputFolder

    ' Instantiate the objects we need
    Set objFileSystem = CreateObject("Scripting.fileSystemObject")
    Set objShell = CreateObject("Shell.Application")

    populateSheetFromFolder = False

    ' The output will now go into the sheet
    If strOutputSheet <> "" Then
        strInputFolder = browseFolder("Owner", False)

        If strInputFolder <> "" Then
            gnCount = getRangeValue("rangeComparingStartRow")
            Call listSubFoldersAndFiles(strInputFolder, strOutputSheet, intHighlightRows)
            populateSheetFromFolder = True
        End If
    End If

End Function

