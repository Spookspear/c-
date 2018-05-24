Attribute VB_Name = "modZapSheet"
Option Explicit

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Public Function ZapWorkSheet(wksWhichSheet As Worksheet, intFirstRow As Currency)
Dim intSelectedRow
    
    wksWhichSheet.Activate
    
    intFirstRow = intFirstRow + 1               ' Delete all rows, bar the header
    Rows(CStr(intFirstRow) & ":" & CStr(intFirstRow)).Select
    
    ' But only if the rows selected are greater than the header
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    
    intSelectedRow = Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Row
    
    If intSelectedRow >= intFirstRow Then
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Rows(Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Address).Select
        Selection.Delete Shift:=xlUp
    End If
    
    Call NavigateHome(wksWhichSheet)
    
End Function

