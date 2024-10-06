VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsresolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// define constants
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1

Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const MIN_SOCKETS_REQD As Long = 1

Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&

Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
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
 
'// private functions
Private Function InitializeSocket() As Boolean
    Dim WSAD As WSADATA
   
    'attempt to initialize the socket
    InitializeSocket = apiWSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

Private Sub CloseSocket()
    'try to close the socket
    If apiWSACleanup() <> 0 Then
        MsgBox "Error calling apiWSACleanup.", vbCritical
    End If

End Sub

Public Function GetIPFromHostName(ByVal sHostName As String) As String
    'converts a host name to an IP address.
   
    Dim nBytes As Long
    Dim ptrHosent As Long
    Dim hstHost As HOSTENT
    Dim ptrName As Long
    Dim ptrAddress As Long
    Dim ptrIPAddress As Long
    Dim sAddress As String 'declare this as Dim sAddress(1) As String if you want 2 ip addresses returned
   
    'try to initalize the socket
    If InitializeSocket() = True Then
      
        'try to get the IP
        ptrHosent = apiGetHostByName(sHostName & vbNullChar)
       
        If ptrHosent <> 0 Then
                   
            'get the IP address
            apiCopyMemory hstHost, ByVal ptrHosent, LenB(hstHost)
            apiCopyMemory ptrIPAddress, ByVal hstHost.hAddrList, 4
             
            'fill buffer
            sAddress = Space$(4)
            'if you want multiple domains returned,
            'fill all items in sAddress array with 4 spaces
           
            apiCopyMemory ByVal sAddress, ByVal ptrIPAddress, hstHost.hLength
           
            'change this to
            'CopyMemory ByVal sAddress(0), ByVal ptrIPAddress, hstHost.hLength
            'if you want an array of ip addresses returned
            '(some domains have more than one ip address associated with it)
           
            'get the IP address
            GetIPFromHostName = IPToText(sAddress)
            'if you are using multiple addresses, you need IPToText(sAddress(0)) & "," & IPToText(sAddress(1))
            'etc
        End If
    Else
        MsgBox "Failed to open Socket."
    End If
End Function

Private Function IPToText(ByVal IPAddress As String) As String
    'converts characters to numbers
    IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function

Public Function GetHostNameFromIP(ByVal sIPAddress As String) As String
   
    Dim ptrHosent As Long
    Dim hAddress As Long
    Dim sHost As String
    Dim nBytes As Long
   
    'try to open the socket
    If InitializeSocket() = True Then
   
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
                    GetHostNameFromIP = sHost
                End If
            Else
                MsgBox "Call to gethostbyaddr failed."
            End If
            'close the socket
            CloseSocket
        Else
            MsgBox "Invalid IP address"
        End If
    Else
        MsgBox "Failed to open Socket"
    End If
End Function

