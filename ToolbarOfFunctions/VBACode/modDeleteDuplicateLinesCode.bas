Attribute VB_Name = "modDeleteDuplicateLinesCode"
Option Explicit

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | dealWithManyDuplicates
' Author    | G Bishop
' Date      | 03/03/2017
' Purpose   | Deals with many columns duplicate only
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
Function dealWithManyDuplicates(WksActiveSheet As Worksheet, intColumToCheck As Integer) As Currency
Dim lngRow As Long
Dim lngCol As Long
Dim strColourDeleteOrClear As String
Dim intNoCheckCols As Integer
Dim strCheckColLtr As String
Dim curNoRecords As Currency
Dim strRange As String

    G_boolDevelopmentMode = True
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Highlight or delete
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strColourDeleteOrClear = getRangeValue("rangeHighlightOrDeleteOption")
    intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck")
    
    strCheckColLtr = getColLetter(intColumToCheck)
    lngRow = getRangeValue("rangeComparingStartRow")
    lngCol = intColumToCheck

    ' clear formatting
    If strColourDeleteOrClear = "Highlight" Then Call clearFormatting
    
    With WksActiveSheet
    
        ' out loop
        Do Until Len(Trim(.Cells(lngRow, intColumToCheck).Value)) = 0
        
            ' check all relevant columns
            For lngCol = intColumToCheck To (intColumToCheck + intNoCheckCols) - 1
                If G_boolDevelopmentMode Then
                    .Cells(lngRow, lngCol).Select
                End If
                
                If .Cells(lngRow, lngCol).Value <> .Cells(lngRow + 1, lngCol).Value Then
                    Exit For
                End If
            Next lngCol
            
            ' if all columns were the same
            If lngCol = (intColumToCheck + intNoCheckCols) Then
                    
                curNoRecords = curNoRecords + 1
                
                If strColourDeleteOrClear = "Delete" Then
                    .Rows(lngRow + 1).Delete
                    lngRow = lngRow - 1
                    
                ElseIf strColourDeleteOrClear = "Highlight" Then
                    strRange = strCheckColLtr & lngRow + 1
                    strRange = strRange & ":" & getColLetter((intColumToCheck + intNoCheckCols) - 1) & lngRow + 1
                
                    .Range(strRange).Font.Color = vbRed
                    
                ElseIf strColourDeleteOrClear = "ClearCell" Then
                    
                     ' .Range("A" & lngRow + 1 & ":" & strCheckColLtr & lngRow + 1).Select
                     .Range("A" & lngRow + 1 & ":" & strCheckColLtr & lngRow + 1).Value = ""
                    lngRow = lngRow + 1 ' jump over new empty line
                End If
            End If
            lngRow = lngRow + 1
        
        Loop
    End With
    
    dealWithManyDuplicates = curNoRecords
    
End Function

'--------------------------------------------------------------------------
' Deals with 1 columns duplicate only
'--------------------------------------------------------------------------
Sub dealWithSingleDuplicates(WksActiveSheet As Worksheet)
Dim lngX As Long
Dim strColourOrDelete As String
Dim intColumToCheck As Integer
Dim lngCount As Long: lngCount = 0
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Highlight or delete
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strColourOrDelete = getRangeValue("rangeHighlightOrDeleteOption")
    intColumToCheck = getRangeValue("rangeDupliateColumnToCheck")

    lngX = 1
    
    Call myDisplayMessage("", "On")
    
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, intColumToCheck).Value)) = 0
            If .Cells(lngX, intColumToCheck).Value = .Cells(lngX + 1, intColumToCheck).Value Then
                Do While .Cells(lngX, intColumToCheck).Value = .Cells(lngX + 1, intColumToCheck).Value
                
                    If strColourOrDelete = "Delete" Then
                        WksActiveSheet.Rows(lngX + 1).Delete
                        lngCount = lngCount + 1
                        
                    ElseIf strColourOrDelete = "Highlight" Then
                        WksActiveSheet.Rows(lngX + 1).Font.Color = vbRed
                        lngX = lngX + 1
                        lngCount = lngCount + 1
                        
                    ElseIf strColourOrDelete = "ClearCell" Then
                        WksActiveSheet.Cells(lngX, 1).Value = ""
                        lngX = lngX + 1
                        lngCount = lngCount + 1
                        
                    End If
                    
                    Call myDisplayMessage("Row: " & CStr(lngX) & " " & strColourOrDelete & "#: " & CStr(lngCount), "Display")

                Loop
            End If

            Call myDisplayMessage("Row: " & CStr(lngX) & " " & strColourOrDelete & "#: From: " & CStr(lngCount), "Display")
            
            lngCount = 0
            lngX = lngX + 1
        Loop
        If lngCount > 0 Then MsgBox "Rows: " & Format(CStr(lngCount), "#0,000"), vbInformation
        
    End With
    
    Call myDisplayMessage("", "Off")
    
End Sub

'--------------------------------------------------------------------------
' inc array on Mode A also - later
'--------------------------------------------------------------------------
Sub delBlankLines(WksActiveSheet As Worksheet, strModeAorB As String)
Dim lngX As Long
Dim lngY As Long
Dim lngCount As Long:  lngCount = 0
Dim lngCountY As Long
ReDim arrRowNos(0)
Dim strAskUser As String
Dim intProceed As Integer: intProceed = vbYes
  
    strAskUser = getRangeValue("rangeProduceMessageBox")
  
    With WksActiveSheet
    
        If strModeAorB = "A" Then
        
            For lngX = .Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
                If WorksheetFunction.CountA(.Rows(lngX)) = 0 Then
                    lngCount = lngCount + 1
                    ReDim Preserve arrRowNos(lngCount)
                    arrRowNos(lngCount) = lngX
                End If
            Next
            
        Else        ' Mode B - Has a warning
        
            For lngX = .Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
                
                lngCountY = .Cells.SpecialCells(xlCellTypeLastCell).Column
                
                For lngY = 1 To lngCountY
                    If Len(allTrim(.Cells(lngX, lngY).Value)) > 0 Then Exit For
                Next lngY
                    
                If lngY = (lngCountY + 1) Then
                    lngCount = lngCount + 1
                    ReDim Preserve arrRowNos(lngCount)
                    arrRowNos(lngCount) = lngX
                End If
                
            Next lngX
        End If
        If lngCount > 0 Then
            If UBound(arrRowNos) > 0 Then
            
                If strAskUser = "Y" Then
                    intProceed = MsgBox("DELETE: There are potentially " & UBound(arrRowNos) & " rows that could be removed on {" & ActiveSheet.Name & "} - proceed?", vbQuestion + vbYesNo)
                End If
                
                If intProceed = vbYes Then
                    For lngX = 1 To UBound(arrRowNos)
                        WksActiveSheet.Rows(arrRowNos(lngX)).Delete
                    Next lngX
                End If
                
            End If
        End If
    End With

End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub highLighManyDuplicates(WksActiveSheet As Worksheet)
Dim lngX As Long
    lngX = 1
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, 1).Value)) = 0
            If Not isSheetEnd(lngX) Then
                If .Cells(lngX, 1).Value = .Cells(lngX + 1, 1).Value Then
                    Do While .Cells(lngX, 1).Value = .Cells(lngX + 1, 1).Value
                        'WksActiveSheet.Rows(lngX + 1).Delete
                        WksActiveSheet.Rows(lngX + 1).Font.Color = vbRed
                    Loop
                End If
                lngX = lngX + 1
            End If
        Loop
    End With
End Sub

