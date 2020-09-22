Attribute VB_Name = "IPAddress"
Public Sub ReptIPAddress()
Dim IP As String
IP = STIP_Address
    'Set up the Winsock control
    With SurfTimeSocket
        .Protocol = sckUDPProtocol
        .LocalPort = BROADCASTPORT
        .RemotePort = BROADCASTPORT
        .RemoteHost = "255.255.255.255" ' This is the broadcast IP
    End With
'Send the IP address of the computer to SurfTime Manager.
'SurfTimeSocket.SendData ST_IPADDRESS
'SurfTimeSocket.SendData IP#
End Sub
Public Function STIP_Address() As String
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim IP As String

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Socket Error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If
    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll error."
        Exit Function
    End If
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
   
    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            IP_Address = IP_Address & temp_ip_address(i) & "."
            cnt = cnt + 1
        Next
        IP_Address = Mid$(IP_Address, 1, Len(IP_Address) - 1)
    Loop While (cnt < 4)
STIP_Address = IP_Address
End Function

Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function


Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function


Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)


    If iReturn <> 0 Then
        MsgBox "Winsock.dll Error."
        End
    End If
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        'sMsg = sMsg & " winsock.dll tarafindan desteklenmiyor. "
        MsgBox sMsg
        End
    End If

End Sub

