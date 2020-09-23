Attribute VB_Name = "mdlwinsock"
Option Explicit

Private Const SIO_RCVALL = &H98000001
Private Const SO_RCVTIMEO = &H1006
Private Const AF_INET = 2
Private Const INVALID_SOCKET = -1
Public Const FD_READ = &H1&
Private Const SOCK_STREAM = 1
Private Const SOCK_RAW = 3
Private Const IPPROTO_IP = 0
Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Private Const INADDR_NONE = &HFFFF
Private Const SOL_SOCKET = &HFFFF&
Private Const hostent_size = 16

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type
Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Public Type ipheader
    ip_verlen As Byte
    ip_tos As Byte
    ip_totallength As Integer
    ip_id As Integer
    ip_offset As Integer
    ip_ttl As Byte
    ip_protocol As Byte
    ip_checksum As Integer
    ip_srcaddr As Long
    ip_destaddr As Long
End Type

Public Type tcpheader
    src_portno As Integer
    dst_portno As Integer
    Sequenceno As Long
    Acknowledgeno As Long
    DataOffset As Byte
    flag As Byte
    Windows As Integer
    checksum As Integer
    UrgentPointer As Integer
End Type


Public Type udpheader
    src_portno As Integer
    dst_portno As Integer
    udp_length As Integer
    udp_checksum As Integer
End Type

Private Const SIO_GET_INTERFACE_LIST = &H4004747F

Private Type sockaddr_gen
   AddressIn As sockaddr
   filler(0 To 7) As Byte
End Type
  
Private Type INTERFACE_INFO
     iiFlags As Long     ' Interface flags
     iiAddress As sockaddr_gen     ' Interface address
     iiBroadcastAddress As sockaddr_gen     ' Broadcast address
     iiNetmask As sockaddr_gen     ' Network mask
End Type
Private Type aINTERFACE_INFO
   interfaceinfo(0 To 7) As INTERFACE_INFO
End Type

Private Declare Function bind Lib "WSOCK32.DLL" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
Private Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function WSAIsBlocking Lib "WSOCK32.DLL" () As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function recv Lib "WSOCK32.DLL" (ByVal s As Long, Buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer
Private Declare Function socket Lib "WSOCK32.DLL" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function closesocket Lib "WSOCK32.DLL" (ByVal s As Long) As Long
Public Declare Function WSAAsyncSelect Lib "WSOCK32.DLL" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname As String, ByVal HostLen As Long) As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Private Declare Function WSACancelBlockingCall Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Private saZero As sockaddr
Private WSAStartedUp As Boolean
Public lSocket As Long


Public Function StartWinsock() As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Public Sub EndWinsock()
    If WSAIsBlocking Then Call WSACancelBlockingCall
    WSACleanup
    WSAStartedUp = False
End Sub

Public Function ConnectSock(ByVal host As String, ByVal Port As Long, ByVal HWndToMsg As Long) As Long ', ByVal Async As Integer) As Long
    
    Dim SockIn           As sockaddr
    Dim lngInBuffer      As Long, _
        lngBytesReturned As Long, _
        lngOutBuffer     As Long, _
        s                As Long, _
        SelectOps        As Long, _
        RCVTIMEO         As Long

    SockIn = saZero
    SockIn.sin_family = AF_INET
    SockIn.sin_port = htons(Port)
    If SockIn.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If

    SockIn.sin_addr = GetHostByNameAlias(host$)

    If SockIn.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If


    s = socket(AF_INET, SOCK_RAW, IPPROTO_IP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If


RCVTIMEO = 5000
If setsockopt(s, SOL_SOCKET, SO_RCVTIMEO, (RCVTIMEO), 4) <> 0 Then
    MsgBox "setsockopt err"
    If s > 0 Then closesocket (s)
    Exit Function
End If

If bind(s, SockIn, Len(SockIn)) <> 0 Then
     If s > 0 Then closesocket (s)
     MsgBox "bind err"
     Exit Function
End If


lngInBuffer = 1
If WSAIoctl(s, SIO_RCVALL, lngInBuffer, Len(lngInBuffer), _
            lngOutBuffer, Len(lngOutBuffer), _
            lngBytesReturned, ByVal 0, ByVal 0) <> 0 Then
    If s > 0 Then closesocket (s)
    MsgBox "WSAIoctl err"
    Exit Function
End If
        
SelectOps = FD_READ 'Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
If WSAAsyncSelect(s, HWndToMsg, WINSOCKMSG, ByVal SelectOps) <> 0 Then
    If s > 0 Then closesocket (s)
    ConnectSock = INVALID_SOCKET
    MsgBox "INVALID_SOCKET"
    Exit Function
End If

ConnectSock = s
End Function

Private Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    Dim phe As Long
    Dim heDestHost As HostEnt
    Dim addrList As Long
    Dim retIP As Long
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, hostent_size
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function

Public Function GetAscIp(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr As Long
    Dim nStr As Long
    Dim retString As String
    retString = String$(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIp = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    CopyMemory ByVal retString, ByVal lpStr, nStr
    retString = Left$(retString, nStr)
    GetAscIp = retString
    If Err Then GetAscIp = "255.255.255.255"
End Function

 

Public Function wsck_enum_interfaces(ByRef str() As String) As Long
    Dim lngBytesReturned      As Long
    Dim NumInterfaces         As Integer
    Dim i                     As Integer
    Dim desc                  As String
    Dim buffer                As aINTERFACE_INFO
    Dim lngSocketDescriptor As Long
    Call StartWinsock
    lngSocketDescriptor = socket(AF_INET, SOCK_STREAM, 0)
    If lngSocketDescriptor = 0 Then
       wsck_enum_interfaces = Err.LastDllError
       Exit Function
    End If
    If WSAIoctl(lngSocketDescriptor, _
        SIO_GET_INTERFACE_LIST, ByVal 0, ByVal 0, _
        buffer, 1024, lngBytesReturned, ByVal 0, ByVal 0) Then
            wsck_enum_interfaces = Err.LastDllError
            Exit Function
    End If
    NumInterfaces = CInt(lngBytesReturned / 76)
    ReDim str(NumInterfaces - 1)
    For i = 0 To NumInterfaces - 1
        str(i) = GetAscIp(buffer.interfaceinfo(i).iiAddress.AddressIn.sin_addr) & ";" & _
                 GetAscIp(buffer.interfaceinfo(i).iiNetmask.AddressIn.sin_addr)
    Next i
    Call closesocket(lngSocketDescriptor)
End Function

Public Function IsWindowsNT5() As Boolean
    Dim typOSInfo As OSVERSIONINFO
    typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)
    Call GetVersionEx(typOSInfo)
    If typOSInfo.dwMajorVersion >= 5 Then IsWindowsNT5 = True
End Function



Public Function HostByName(name As String) As String
 Dim hostname As String * 256
   Dim hostent_addr As Long
   Dim host As HostEnt
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String

       If gethostname(hostname, 256) = -1 Then
           MsgBox "Windows Sockets error " & WSAGetLastError()
           
           GetHostByNameAlias name
           Exit Function
       Else
           hostname = Trim$(hostname)
       End If
       If Len(name) > 0 Then Mid(hostname, 1, Len(name)) = name
       hostent_addr = gethostbyname(hostname)

       If hostent_addr = 0 Then
           'MsgBox "Winsock.dll is not responding."
           HostByName = "Unknown"
           Exit Function
       End If

       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.h_addr_list, 4

       Do
           ReDim temp_ip_address(1 To host.h_length)
           RtlMoveMemory temp_ip_address(1), hostip_addr, host.h_length

           For i = 1 To host.h_length
               ip_address = ip_address & temp_ip_address(i) & "."
           Next
           ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

           HostByName = ip_address

           ip_address = ""
           host.h_addr_list = host.h_addr_list + LenB(host.h_addr_list)
           RtlMoveMemory hostip_addr, host.h_addr_list, 4
        Loop While (hostip_addr <> 0)

        

End Function
