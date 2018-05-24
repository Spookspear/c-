Attribute VB_Name = "modSpecial"
Option Explicit
Dim objFileSystem

Sub nowRemoveFiles()
Dim WksActiveSheet As Worksheet
'Dim intWalkVal As Integer
Dim strSting1 As String
Dim lngX As Long
Dim intDelCount As Integer

    Set objFileSystem = CreateObject("Scripting.fileSystemObject")
    Set WksActiveSheet = ActiveSheet

    lngX = 2
    intDelCount = 2
    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngX, 1).Value)) = 0
        
            strSting1 = .Cells(lngX, 1).Value
            
            If .Cells(lngX, 1).Font.Color = vbRed Then
                Call killFile(strSting1)
                intDelCount = intDelCount + 1
            End If
            
            lngX = lngX + 1

        Loop
        
    End With
    MsgBox "Number of files deleted..: " & CStr(intDelCount)
    Set objFileSystem = Nothing
    Set WksActiveSheet = Nothing

End Sub

' ----------+-----------------------------------------------------------------------------------------------------
' Procedure | keepOneRemoveOne
' Author    | G Bishop
' Date      | 14/09/2016
' Purpose   | Put the line underneath next door and skip 2
' ----------+------------+----------------------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+----------------------------------------------------------------------------------------
' G Bishop  | 14/09/2016 | Added in this header
' ----------+------------+----------------------------------------------------------------------------------------
Sub keepOneRemoveOne()
Dim WksActiveSheet As Worksheet
Dim lngRow As Long

    Set WksActiveSheet = ActiveSheet

    lngRow = 1

    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngRow, 1).Value)) = 0
            .Cells(lngRow, 2).Value = .Cells(lngRow + 1, 1).Value
            .Cells(lngRow + 1, 1).Value = ""
            
            lngRow = lngRow + 2
        
        Loop
    End With

End Sub


Sub SHOWALL()
Dim WksActiveSheet As Worksheet
Dim lngRow As Long
Dim intSpace As Integer
Dim strSpace
Dim strValue1 As String

    Set WksActiveSheet = ActiveSheet

    intSpace = 41
    strSpace = Space(intSpace)

    lngRow = 1
    ' intDelCount = 2

    With WksActiveSheet
        Do Until Len(Trim(.Cells(lngRow, 1).Value)) = 0
            strValue1 = Trim(.Cells(lngRow, 1).Value)
            
            intSpace = 41 - Len(strValue1)
            Debug.Print strValue1 & Space(intSpace) & Trim(.Cells(lngRow, 2).Value)
            lngRow = lngRow + 1
        
        Loop
    End With

End Sub

' --------------------------------------------------------------------------------------
' Procedure | trimOutFromBraces
' Author    | Grant Bishop
' Date      | 01/05/2018
' Purpose   | extracts values from between parentheses
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub trimOutFromBraces()
Dim Wks As Worksheet
Dim lngRow As Long
Dim strValue As String

Dim intStart As Integer
Dim intEnd As Integer
Dim strDiscard As String



    Set Wks = ActiveSheet
    
    lngRow = 2
    
    With Wks
    
        ' out loop
        Do Until Len(Trim(.Cells(lngRow, 1).Value)) = 0
           
            
            Debug.Print .Cells(lngRow, 1).Value
            
            strValue = .Cells(lngRow, 1).Value
            
            If InStr(strValue, "(") > 0 Then
            
                If InStr(strValue, ")") > 0 Then
                
                    intStart = InStr(strValue, "(") + 1
                    intEnd = InStr(strValue, ")")
                    
                    Debug.Print intStart
                    Debug.Print intEnd
                    
                    
                    strDiscard = Mid(strValue, 1, intStart - 2)
                    
                    Debug.Print Mid(strValue, (intStart), (intEnd - intStart))
                    
                    .Cells(lngRow, 2).Value = Mid(strValue, (intStart), (intEnd - intStart))
                    .Cells(lngRow, 3).Value = strDiscard
                    
                
                
                
                Else
                    MsgBox "non matching ()", , "1GVB1 30/04/2018"
                    
                
                End If
            
            
            End If
            
            
            lngRow = lngRow + 1
            
        
        Loop
    End With
    


End Sub

