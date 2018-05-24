Attribute VB_Name = "modImportADGroup"
Option Explicit

'----------------------------------------------------------------------------------------
' Populate Spread sheet with all the members of a supplied AD group
'----------------------------------------------------------------------------------------
Public Sub LoadADGroupIntoSpreadsheet(wksWhichSheet, strDomainName, objGroupName)
Dim objGroup As Object
Dim objMember As Object
Dim intRow As Currency
Dim strChopName As String
'Dim objDomainNT As Object
    intRow = 2
    
    'On Error GoTo MyErrorHandler
    'On Error Resume Next
    
    Set objGroup = GetObject("WinNT://" & strDomainName & "/" & objGroupName)

    For Each objMember In objGroup.Members
        If objMember.Class = "User" Then
            With objMember

                'Debug.Print .Name
                
                strChopName = UCase(Mid$(.Name, 1, 3))
                
                If strChopName <> "S-1" Then
                
                    If strChopName <> "SS7" Then
                        wksWhichSheet.Cells(intRow, 1) = UCase(.Name)
                        wksWhichSheet.Cells(intRow, 2) = .FullName
                        wksWhichSheet.Cells(intRow, 3) = .Description
                        wksWhichSheet.Cells(intRow, 4) = .AccountDisabled
                        
                        If InStr(.Name, "$") > 0 Then
                            wksWhichSheet.Cells(intRow, 5) = retrieveValue(getDN(.Name, CStr(strDomainName)), "OU", 3)
                        End If
                    Else
                        wksWhichSheet.Cells(intRow, 1) = UCase(.Name)
                        wksWhichSheet.Cells(intRow, 2) = getAdDetails(.Name, "subsea7.net", "FullName")
                        wksWhichSheet.Cells(intRow, 3) = "'" & getAdDetails(.Name, "subsea7.net", "Description")
                        wksWhichSheet.Cells(intRow, 4) = getAdDetails(.Name, "subsea7.net", "AccountDisabled")
                    
                    End If
                Else
                    wksWhichSheet.Cells(intRow, 1) = UCase(.Name)
                    wksWhichSheet.Cells(intRow, 2) = "The handle is invalid"
                End If

                intRow = intRow + 1
                
            End With
        ElseIf objMember.Class = "objGroup" Then
            MsgBox "Unhandled", vbCritical
        End If

    Next
    
    On Error GoTo 0

Exit Sub
    
MyErrorHandler:
Dim strErrDescription
    strErrDescription = ""
    strErrDescription = strErrDescription & Err.Number & "  -  " & Err.Description & vbNewLine & vbNewLine
    strErrDescription = strErrDescription & "Does the GROUP '" & wksWhichSheet.Name & "' exist?"
    
    MsgBox strErrDescription, vbCritical
   
End Sub


