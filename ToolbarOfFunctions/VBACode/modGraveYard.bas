Attribute VB_Name = "modGraveYard"
Option Explicit

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub ColourRow2(strDoWhat As String, intRow As Integer)
Dim dblWhichFore As Double
Dim dblWhichBack As Double

    Select Case strDoWhat
    Case "New"
        dblWhichFore = InternalParameters.Range("RangeColourNewFore").Interior.ColorIndex
        dblWhichBack = InternalParameters.Range("RangeColourNewBack").Interior.ColorIndex
    Case "Updated"
        dblWhichFore = InternalParameters.Range("RangeColourUpdateFore").Interior.ColorIndex
        dblWhichBack = InternalParameters.Range("RangeColourUpdateBack").Interior.ColorIndex
    Case "NoChange"
        dblWhichFore = InternalParameters.Range("RangeColourPreviousFore").Interior.ColorIndex
        dblWhichBack = InternalParameters.Range("RangeColourPreviousBack").Interior.ColorIndex
    End Select

    Rows(intRow).Select
    
    With Selection.Interior
        .ColorIndex = dblWhichBack
        Selection.Font.ColorIndex = dblWhichFore
    End With

End Sub

'--------------------------------------------------------------------------
' Deals with many columns duplicate only
'--------------------------------------------------------------------------
Sub dealWithManyDuplicatesOld(WksActiveSheet As Worksheet)
Dim lngX As Long
Dim lngY As Long
Dim strColourOrDelete As String
Dim intNoCheckCols As Integer
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Highlight or delete
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strColourOrDelete = getRangeValue("rangeHighlightOrDeleteOption")
    intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck") + 1
    lngX = 1
    lngY = 1

    ' clear formatting
    If strColourOrDelete = "Highlight" Then Call clearFormatting
    
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, 1).Value)) = 0
            If Not isSheetEnd(lngX) Then
                Do Until lngY = intNoCheckCols
                    If .Cells(lngX, lngY).Value <> .Cells(lngX + 1, lngY).Value Then
                        Exit Do
                    End If
                    lngY = lngY + 1
                Loop
                
                If lngY = intNoCheckCols Then   ' can delete one row only
                    If strColourOrDelete = "Delete" Then
                        WksActiveSheet.Rows(lngX + 1).Delete
                        lngX = lngX - 1
                    Else
                        WksActiveSheet.Rows(lngX + 1).Font.Color = vbRed
                    End If
                End If
                lngX = lngX + 1: lngY = 1
            End If
        Loop
    End With
End Sub

'--------------------------------------------------------------------------
' Deals with many columns duplicate only
'--------------------------------------------------------------------------
Sub dealWithManyDuplicatesOld1(WksActiveSheet As Worksheet)
Dim lngX As Long
Dim lngY As Long
Dim strColourOrDelete As String
Dim intNoCheckCols As Integer
    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Highlight or delete
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strColourOrDelete = getRangeValue("rangeHighlightOrDeleteOption")
    intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck") + 1
    lngX = 1
    lngY = 1

    ' clear formatting
    If strColourOrDelete = "Highlight" Then Call clearFormatting
    
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, 1).Value)) = 0
            If Not isSheetEnd(lngX) Then
                Do Until lngY = intNoCheckCols
                    If .Cells(lngX, lngY).Value <> .Cells(lngX + 1, lngY).Value Then
                        Exit Do
                    End If
                    lngY = lngY + 1
                Loop
                
                If lngY = intNoCheckCols Then   ' can delete one row only
                    If strColourOrDelete = "Delete" Then
                        WksActiveSheet.Rows(lngX + 1).Delete
                        lngX = lngX - 1
                    Else
                        WksActiveSheet.Rows(lngX + 1).Font.Color = vbRed
                    End If
                End If
                lngX = lngX + 1: lngY = 1
            End If
        Loop
    End With
End Sub

'--------------------------------------------------------------------------
' Deals with many columns duplicate only
'--------------------------------------------------------------------------
Function dealWithManyDuplicatesOld2(WksActiveSheet As Worksheet, intColumToCheck As Integer) As Currency
Dim lngX As Long
Dim lngY As Long
Dim strColourDeleteOrClear As String
Dim intNoCheckCols As Integer
Dim strCheckColLtr As String
Dim curNoDeletions As Currency

    
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    ' Highlight or delete
    '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    strColourDeleteOrClear = getRangeValue("rangeHighlightOrDeleteOption")
    intNoCheckCols = getRangeValue("rangeNoOfColumnsToCheck")
    strCheckColLtr = getColLetter(intNoCheckCols)
    intNoCheckCols = intNoCheckCols + 1
    lngX = 1
    lngY = 1

    ' clear formatting
    If strColourDeleteOrClear = "Highlight" Then Call clearFormatting
    
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, intColumToCheck).Value)) = 0
            Do Until lngY = intNoCheckCols
                If .Cells(lngX, lngY).Value <> .Cells(lngX + 1, lngY).Value Then
                    Exit Do
                End If
                lngY = lngY + 1
            Loop
            
            If lngY = intNoCheckCols Then   ' not the same = found a difference
                    
                curNoDeletions = curNoDeletions + 1
                
                If strColourDeleteOrClear = "Delete" Then
                    .Rows(lngX + 1).Delete
                    lngX = lngX - 1
                    
                ElseIf strColourDeleteOrClear = "Highlight" Then
                    .Range("A" & lngX + 1 & ":" & strCheckColLtr & lngX + 1).Font.Color = vbRed
                    
                ElseIf strColourDeleteOrClear = "ClearCell" Then
                     ' .Range("A" & lngX + 1 & ":" & strCheckColLtr & lngX + 1).Select
                     .Range("A" & lngX + 1 & ":" & strCheckColLtr & lngX + 1).Value = ""
                    lngX = lngX + 1 ' jump over new empty line
                End If
            End If
            lngX = lngX + 1: lngY = 1
        Loop
    End With
    
    dealWithManyDuplicatesOld2 = curNoDeletions
    
End Function

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub DelBlankLinesOld(WksActiveSheet As Worksheet, strModeAorB As String)
Dim lngX As Long
Dim lngY As Long
Dim lngCount As Long:  lngCount = 0
Dim lngCountY As Long
ReDim arrRowNos(0)
'Dim lProceed As Boolean

    With WksActiveSheet
    
        If strModeAorB = "A" Then
        
            For lngX = .Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
                If WorksheetFunction.CountA(.Rows(lngX)) = 0 Then
                    WksActiveSheet.Rows(lngX).Delete
                    lngCount = lngCount + 1
                End If
            Next
            
        Else        ' Mode B - Has a warning
        
            For lngX = .Cells.SpecialCells(xlCellTypeLastCell).Row To 1 Step -1
                
                lngCountY = .Cells.SpecialCells(xlCellTypeLastCell).Column
                
                For lngY = 1 To lngCountY
                    If Len(Trim(.Cells(lngX, lngY).Value)) > 0 Then Exit For
                Next lngY
                    
                If lngY = (lngCountY + 1) Then
                    lngCount = lngCount + 1
                    ReDim Preserve arrRowNos(lngCount)
                    arrRowNos(lngCount) = lngX
                End If
                
            Next lngX
        End If
        If lngCount > 0 Then
            If strModeAorB = "A" Then
                If getRangeValue("RangeProduceMessageBox") = "Y" Then
                    MsgBox "Rows removed: " & CStr(lngCount), vbInformation
                End If
            Else
                If UBound(arrRowNos) > 0 Then
                    If MsgBox("DELETE: There are potentially " & UBound(arrRowNos) & " rows that could be removed on {" & ActiveSheet.Name & "} - proceed?", vbQuestion + vbYesNo) = vbYes Then
                    'If lProceed Then
                        For lngX = 1 To UBound(arrRowNos)
                            WksActiveSheet.Rows(arrRowNos(lngX)).Delete
                        Next lngX
                    End If
                End If
            
            End If
        End If
    End With
    
    ' ActiveWorkbook.Save

End Sub

