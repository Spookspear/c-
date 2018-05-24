Attribute VB_Name = "modImportADUser"
Option Explicit

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub readGroupMembership(wksWhichSheet As Worksheet)
Dim strUserName As String
Dim strDomain As String
Dim intRow As Currency
Dim dsRoot As String
Dim dsobj As Object
Dim Prop As Object

    Set wksWhichSheet = wksWhichSheet
    strUserName = ActiveSheet.Name
    
    strDomain = GrabDomainOrGroup(ActiveSheet.Name, "Domain", "@")
    strUserName = GrabDomainOrGroup(ActiveSheet.Name, "User", "@")
    
    intRow = 2
    
    If Len(Trim(strUserName)) > 0 Then
        
        dsRoot = "WinNT://" & strDomain & "/" & strUserName
        
        'On Error GoTo MyErrorHandler
        Set dsobj = GetObject(dsRoot)
        
        For Each Prop In dsobj.Groups
            wksWhichSheet.Cells(intRow, 1).Value = Prop.Name
            wksWhichSheet.Cells(intRow, 2).Value = Prop.Description
            intRow = intRow + 1
        Next
        On Error GoTo 0
        
    End If
        
Exit Sub

MyErrorHandler:
Dim strErrDescription
    strErrDescription = ""
    strErrDescription = strErrDescription & Err.Number & "  -  " & Err.Description & vbNewLine & vbNewLine
    strErrDescription = strErrDescription & "Is there are Active Directory USER called: '" & wksWhichSheet.Name & "' ?"
    MsgBox strErrDescription, vbCritical
End Sub

' --------------------------------------------------------------------------------------
' Procedure | readADUserDetails
' Author    | Grant Bishop
' Date      | 14/05/2018
' Purpose   | Reads User Names and populates sheet with Full name and if active
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub readADUserDetails(WksActiveSheet As Worksheet)
Dim intRow As Integer
Dim strDomain As String
Dim strUserName As String

    ' default
    strDomain = "GROUP"
    
    intRow = 2

    With WksActiveSheet
        Do Until Len(Trim(.Cells(intRow, 1).Value)) = 0
            Debug.Print .Cells(intRow, 1).Value
            strUserName = .Cells(intRow, 1).Value
            
            strDomain = GrabDomainOrGroup(strUserName, "Domain", "\")
            strUserName = GrabDomainOrGroup(strUserName, "User", "\")
            
            If Len(Trim(strUserName)) > 0 Then
            
                If Len(Trim(.Cells(intRow, 2).Value)) = 0 Then
                    .Cells(intRow, 2).Value = getAdDetails(strUserName, strDomain, "FullName")
                End If
                
                If Len(Trim(.Cells(intRow, 3).Value)) = 0 Then
                    .Cells(intRow, 3).Value = getAdDetails(strUserName, strDomain, "AccountDisabled")
                End If
                
            End If
            
            intRow = intRow + 1
            
        Loop
        

    End With

End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub Test()

    'Debug.Print GetAdDetails("GBI01", "GROUP")
    'Debug.Print GetAdDetails("SS7U3501", "subsea7.net", "FullName")
    'Debug.Print GetAdDetails("SS7U3501", "subsea7.net", "Description")
    'Debug.Print GetAdDetails("GBI01", GetRangeValue("rangeDomain"), "FullName")
    'Debug.Print GetAdDetails("GBI01", GetRangeValue("rangeDomain"), "Description")
    ' Debug.Print GetAdDetails("GBI01", "GROUP", "FullName")
    'Debug.Print getAdDetails("GBI01", "subsea7.net", "FullName")
    'Debug.Print getAdDetails("robertcgetronics", "S7", "FullName")
    'Debug.Print getAdDetails("Gbishop-adm", "S7", "FullName")
    'Debug.Print GetAdDetails("Gbishop-adm", GetRangeValue("rangeDomain"), "FullName")
    'Debug.Print GetDN("Gbishop-adm")

    Debug.Print getAdDetails("GBI01", "GROUP", "FullName")


End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Sub Test2()
    'Debug.Print GetDN("gbi01-adm", GetRangeValue("rangeDomain"))
    Debug.Print getDN("Gbishop-adm", "S7")
    Debug.Print getDN("Gbishop-adm", "subsea7.net")
    
End Sub

'--------------------------------------------------------------------------
'
'--------------------------------------------------------------------------
Function getAdDetails(strUserName As String, strDomain As String, strGetWhat As String) As String
Dim dsRoot As String
Dim strRetValue As String
Dim dsobj As Object
    
    If Len(Trim(strUserName)) > 0 Then
        
        On Error GoTo MyErrorHandler
        dsRoot = "WinNT://" & strDomain & "/" & strUserName
        
        Set dsobj = GetObject(dsRoot)
        If strGetWhat = "FullName" Then
            strRetValue = dsobj.FullName
        ElseIf strGetWhat = "Description" Then
            strRetValue = dsobj.Description
        ElseIf strGetWhat = "AccountDisabled" Then
            strRetValue = dsobj.AccountDisabled
        End If
        
    End If
        
    On Error GoTo 0
    getAdDetails = strRetValue
    
    Set dsobj = Nothing
Exit Function
MyErrorHandler:
Dim strErrDescription
    strErrDescription = ""
    strErrDescription = strErrDescription & Err.Number & "  -  " & Err.Description & vbNewLine & vbNewLine
    strErrDescription = strErrDescription & "Does the User / domain =' " & strUserName & " / " & strDomain & "' exist?"
    
    MsgBox strErrDescription, vbCritical
    
    
End Function



