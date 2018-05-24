Attribute VB_Name = "modPinger"
Option Explicit
' --------------------------------------------------------------------------------------
'    Author | Grant Bishop
'      Date | 01/05/2018
'      Name | modPinger
'   Purpose | To allow pinging of list of servers in specific column and return results to another column
'     To Do |
' ----------+---------------------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
'// api
'kernel32
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nBytes As Long)
Private Declare Function apiStrLen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
'wsock32
Private Declare Function apiGetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal hostname As String) As Long
Private Declare Function apiWSAStartup Lib "wsock32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function apiWSACleanup Lib "wsock32.dll" Alias "WSACleanup" () As Long
Private Declare Function apiInetAddr Lib "wsock32.dll" Alias "inet_addr" (ByVal s As String) As Long
Private Declare Function apiGetHostByAddr Lib "wsock32.dll" Alias "gethostbyaddr" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long

'// define constants
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1

Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
'Private Const MIN_SOCKETS_REQD As Long = 1

Private Const WS_VERSION_REQD As Long = &H101
'Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
'Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&

'Private Const WSADescription_Len As Long = 256
'Private Const WSASYS_Status_Len As Long = 128
Private Const AF_INET As Long = 2

'// structures
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
 
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

' --------------------------------------------------------------------------------------
' Procedure | addPingSheetToToolbar
' Author    | Grant Bishop
' Date      | 03/05/2018
' Purpose   | adds the Ping option to the toolbar
' To Do     |
' ----------+------------+--------------------------------------------------------------
' Modified  | Date       | Reason
' ----------+------------+--------------------------------------------------------------
Sub addPingSheetToToolbar(ByRef cbMainMenuBar As CommandBar, ByRef intButtonCount, intStyle As Currency, boolInclude As Boolean)
Dim myNewButton As CommandBarButton

    If boolInclude Then
    
        intButtonCount = intButtonCount + 1
        Set myNewButton = cbMainMenuBar.Controls.Add(Type:=msoControlButton, Before:=intButtonCount)
        With myNewButton
            .Caption = "&Ping Servers"
            .TooltipText = "will ping all servers in this sheet"
            .Style = intStyle
            .FaceId = 700
            .OnAction = "btnPingServers"
        End With
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------
' Procedure : closeSocket
' Author    : the Internet
' Date      : 06/10/2015
' Purpose   : try to close the socket
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Private Sub closeSocket()
    If apiWSACleanup() <> 0 Then
        Call MsgBox("Error calling apiWSACleanup", vbCritical Or vbSystemModal Or vbDefaultButton1, GC_MSGBOX)
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------
' Procedure : getHostNameFromIP
' Author    : Grant Bishop
' Date      : 06/10/2015
' Purpose   : try to open the socket, convert string address to long datatype
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Public Function getHostNameFromIP(ByVal sIPAddress As String) As String
Dim ptrHosent As Long
Dim hAddress As Long
Dim sHost As String
Dim nBytes As Long
    
    'try to open the socket
    If initializeSocket() = True Then
    
        'convert string address to long datatype
        hAddress = apiInetAddr(sIPAddress)
        
        'check if an error ocucred
        If hAddress <> SOCKET_ERROR Then
            
            'obtain a pointer to the HOSTENT structure
            'that contains the name and address
            'corresponding to the given network address.
            ptrHosent = apiGetHostByAddr(hAddress, 4, AF_INET)
            
            If ptrHosent <> 0 Then
                
                'convert address and
                'get resolved hostname

                apiCopyMemory ptrHosent, ByVal ptrHosent, 4
                
                nBytes = apiStrLen(ByVal ptrHosent)
                
                If nBytes > 0 Then
                    'fill the IP address buffer
                    sHost = Space$(nBytes)
                    
                    apiCopyMemory ByVal sHost, ByVal ptrHosent, nBytes
                    getHostNameFromIP = sHost
                End If
            Else
                Call MsgBox("Call to gethostbyaddr failed.", vbCritical Or vbSystemModal Or vbDefaultButton1, GC_MSGBOX)
                
            End If
            'close the socket
            closeSocket
        Else
            Call MsgBox("Invalid IP address", vbExclamation Or vbSystemModal Or vbDefaultButton1, GC_MSGBOX)
            
        End If
    Else
        Call MsgBox("Failed to open Socket", vbExclamation Or vbSystemModal Or vbDefaultButton1, GC_MSGBOX)
    End If
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : getIPFromHostName
' Author    : Grant Bishop
' Date      : 06/10/2015
' Purpose   : converts a host name to an IP address.
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Public Function getIPFromHostName(ByVal sHostName As String) As String
Dim ptrHosent As Long
Dim hstHost As HOSTENT
Dim ptrIPAddress As Long
Dim sAddress As String              'declare this as Dim sAddress(1) As String if you want 2 ip addresses returned
    
    'try to initalize the socket
    If initializeSocket() = True Then
       
        'try to get the IP
        ptrHosent = apiGetHostByName(sHostName & vbNullChar)
        
        If ptrHosent <> 0 Then
                    
            'get the IP address
            Call apiCopyMemory(hstHost, ByVal ptrHosent, LenB(hstHost))
            Call apiCopyMemory(ptrIPAddress, ByVal hstHost.hAddrList, 4)
              
            'fill buffer
            sAddress = Space$(4)
            'if you want multiple domains returned,
            'fill all items in sAddress array with 4 spaces
            
            Call apiCopyMemory(ByVal sAddress, ByVal ptrIPAddress, hstHost.hLength)
            
            'change this to
            'CopyMemory ByVal sAddress(0), ByVal ptrIPAddress, hstHost.hLength
            'if you want an array of ip addresses returned
            '(some domains have more than one ip address associated with it)
            
            'get the IP address
            getIPFromHostName = IPToText(sAddress)
            'if you are using multiple addresses, you need IPToText(sAddress(0)) & "," & IPToText(sAddress(1))
            'etc
        End If
    Else
        Call MsgBox("Failed to open Socket.", vbExclamation Or vbSystemModal, GC_MSGBOX)
    End If
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : initializeSocket
' Author    : Grant Bishop
' Date      : 06/10/2015
' Returns   : Boolean bool
' Purpose   : private functions attempt to initialize the socket
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Private Function initializeSocket() As Boolean
Dim WSAD As WSADATA
    initializeSocket = apiWSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : IPToText
' Author    : Grant Bishop
' Date      : 06/10/2015
' Returns   : String str
' Purpose   : converts characters to numbers
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Private Function IPToText(ByVal IPAddress As String) As String
    IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : meetsConditions
' Author    : Grant Bishop
' Date      : 06/10/2015
' Returns   : Boolean bool
' Purpose   : used to test specific cells on Wks.intRow, can be expanded to add more rules if required
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Function meetsConditions(Wks As Worksheet, intRow As Integer, intPingColWrite As Integer) As Boolean
    meetsConditions = (Len(Trim(Wks.Cells(intRow, intPingColWrite).Value)) = 0)
        
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : pingerDosPing
' Author    : Grant Bishop
' Date      : 06/10/2015
' Returns   : Integer int
' Purpose   : Uses Object File System to create a DOS window and grab results using StdOut.ReadAll
'             chops up message and looks for known text
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Function pingerDosPing(strIPAddress) As Integer
Dim objShell
Dim objScriptExec
Dim strPingResults
Dim strPingRate
Dim intPingRate
Dim intCutPos1
Dim intCutPos2

    Set objShell = CreateObject("WScript.Shell")
    Set objScriptExec = objShell.Exec("ping -n 2 -w 3000 " & strIPAddress)
    
    strPingResults = LCase(objScriptExec.StdOut.ReadAll)

    If InStr(strPingResults, "reply from") Then
        If InStr(strPingResults, "destination net unreachable") Then
            intPingRate = -1
            
        ElseIf InStr(strPingResults, "ttl expired in transit") Then
            intPingRate = -1
        
        ElseIf InStr(strPingResults, "destination host unreachable") Then
            intPingRate = -1
        
        Else
            
            ' All ok - try and get the ping rate
            If G_boolDevelopmentMode Then Debug.Print Right(strPingResults, 10)
            
            strPingRate = Trim(Right(strPingResults, 10))
            
            If G_boolDevelopmentMode Then Debug.Print strPingRate
            
            intCutPos1 = InStr(strPingRate, "=") + 1
            intCutPos2 = InStr(strPingRate, "ms") - (1 + (intCutPos1 - 1))
            
            intPingRate = Trim(Mid(strPingRate, intCutPos1, intCutPos2))
            
        End If
    Else
        intPingRate = -1
    End If
    
    pingerDosPing = intPingRate
    
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : pingProcessSheet
' Author    : Grant Bishop
' Date      : 06/10/2015
' Returns   : Worksheet obj
' Purpose   : received a worksheet as a parameter
'----------------------------------------------------------------------------------------------------------------
' Modified by   | Date      | Reason
'----------------------------------------------------------------------------------------------------------------
Function pingProcessSheet(Wks As Worksheet) As Worksheet
Dim strHost As String
Dim intRow As Integer
Dim intPingRate As Integer
Dim strHostName As String

Dim intPingColRead As Integer
Dim intPingColWrite As Integer

    G_boolDevelopmentMode = False

    intRow = 2
    intRow = getRangeValue("rangePingSheetRowNo")
    
    ' Find columns to work on
    'G_intVessel = searchForValueInCol(Wks, "Remote Location", 1, "Whole")
    'intPingColRead = searchForValueInCol(Wks, "Server", 1, "Whole")
    'intPingColWrite = searchForValueInCol(Wks, "Rate", 1, "Whole")
        
    intPingColRead = getRangeValue("rangeColPingRead")
    intPingColWrite = getRangeValue("rangeColPingWrite")
    
    With Wks
        If G_boolDevelopmentMode Then .Activate
    
        Do Until Len(Trim(.Cells(intRow, intPingColRead).Value)) = 0
        
            ' DEBUG
            If G_boolDevelopmentMode Then
                If intRow = 12 Then
                    MsgBox "Stop", , "1gvb2"
                End If
            End If
            
            ' DEBUG
        
            If meetsConditions(Wks, intRow, intPingColWrite) Then
            
                If G_boolDevelopmentMode Then .Cells(intRow, intPingColRead).Select
                
                strHostName = .Cells(intRow, intPingColRead).Value
            
                strHost = getIPFromHostName(strHostName)
                
                If Len(Trim(strHost)) <> 0 Then
                
                    Call formatCells(Wks, intRow, intPingColWrite, "SmallCentred")
                    
                    Application.ScreenUpdating = True
                    .Cells(intRow, intPingColWrite).Value = "Pinging"
                    Application.ScreenUpdating = False
                    intPingRate = pingerDosPing(strHost)
                    
                    'intPingRate = rotateText(Wks, intRow, intPingColWrite, strHost)
                    
                    With .Cells(intRow, intPingColWrite)
                    
                        If G_boolDevelopmentMode Then .Select
                            
                        If intPingRate >= 0 Then
                            .Value = intPingRate
                            Call formatCells(Wks, intRow, intPingColWrite, "Normal")
                        Else
                            .Value = 0
                            Call formatCells(Wks, intRow, intPingColWrite, "Accent2")
                        End If
                    
                    End With

                    ActiveWorkbook.RefreshAll
                Else
                    .Cells(intRow, intPingColWrite).Value = "ERR"
                    Call formatCells(Wks, intRow, intPingColWrite, "Bad")
            
                End If
                
            End If
            
            DoEvents
        
            intRow = intRow + 1
            
        Loop
        
    End With
    
    Set pingProcessSheet = Wks
    
End Function

'----------------------------------------------------------------------------------------------------------------
' Procedure : btnPingServers
' Author    : Grant Bishop
' Date      : 01-05-2018
' Returns   : n/a
' Purpose   : calls pingProcessSheet( Worksheet / active sheet )
'----------------------------------------------------------------------------------------------------------------
Sub pingServers()
Dim Wks As Worksheet
Dim strCellAddress As String
    
    Application.ScreenUpdating = False
    strCellAddress = ActiveCell.Address
    
    Set Wks = pingProcessSheet(ActiveSheet)
    'Set Wks = ActiveSheet
    
    'If MsgBox("Sort results", vbInformation + vbYesNo, GC_MSGBOX) = vbYes Then
    '    Call sortSheet_Pinger(Wks)
    'End If
    
    Wks.Range(strCellAddress).Select

End Sub

