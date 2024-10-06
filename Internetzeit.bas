Attribute VB_Name = "InternetZeit"
Option Explicit
Dim RegPos$
' Fügen Sie diesen Code in ein Öffentliches Modul ein
Private Declare Function gethostbyname& Lib "wsock32.dll" (ByVal name$)
Private Declare Function socket& Lib "wsock32.dll" ( _
  ByVal af&, _
  ByVal prototype&, _
  ByVal protocol&)
Private Declare Function closesocket& Lib "wsock32.dll" (ByVal s&)
Private Declare Function connect& Lib "wsock32.dll" ( _
  ByVal s&, _
  name As SOCKADDR, _
  ByVal namelen&)
Private Declare Function send& Lib "wsock32.dll" ( _
  ByVal s&, _
  buf As Any, _
  ByVal length&, _
  ByVal flags&)
Private Declare Function recv& Lib "wsock32.dll" ( _
  ByVal s&, _
  buf As Any, _
  ByVal length&, _
  ByVal flags&)
Private Declare Function ioctlsocket& Lib "wsock32.dll" ( _
  ByVal s&, _
   ByVal cmd&, _
  argp&)
Private Declare Function inet_addr& Lib "wsock32.dll" ( _
  ByVal cp$)
Private Declare Function htons% Lib "wsock32.dll" ( _
  ByVal hostshort%)
Private Declare Function WSAGetLastError& Lib "wsock32.dll" ()
Private Declare Sub MoveMemory Lib "kernel32" _
  Alias "RtlMoveMemory" ( _
  Destination As Any, _
  source As Any, _
  ByVal length&)
    
Private Type HOSTENT
  hname As Long
  haliases As Long
  haddrtype As Integer
  hlength As Integer
  haddrlist As Long
End Type
 
Private Type SOCKADDR
  sin_family As Integer
  sin_port As Integer
  sin_addr As Long
  sin_zero As String * 8
End Type
 
' eine der HOSTENT Hardtype-Konstanten
Private Const AF_INET = 2 ' Internet Protokoll (UDP/IP oder TCP/IP).
 
' Socket Prototype-Konstanten
Private Const SOCK_STREAM = 1 '  2-wege Stream. Bei AF_INET ist es das TCP/IP Protokoll
Private Const SOCK_DGRAM = 2 ' Datagramm basierende Verbindung. Bei AF_INET ist es das UDP Protokoll
 
' recv flags-Konstanten
Private Const MSG_PEEK = &H2 ' Daten aus dem Puffer lesen, aber nicht aus dem Puffer entfernen
 
' ioctlsocket cmd-Konstanten
Private Const FIONBIO = &H8004667E ' Setzen, ob die Funktion bei der nächsten Datenanfrage zurückkehren soll

Private Declare Function WSAStartup& Lib "wsock32.dll" ( _
  ByVal wVersionRequested%, _
  lpWSAData As WSAData)

 
Private Type WSAData
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * 257
  szSystemStatus As String * 129
  iMaxSockets As Long
  iMaxUdpDg As Long
  lpVendorInfo As Long
End Type

Private Declare Function WSACleanup& Lib "wsock32.dll" ()
Dim hSock&

Private Declare Function GetTickCount& Lib "kernel32" ()
Private TimeDelay!, T0!, T1!

Const TIME_ZONE_ID_DAYLIGHT& = 2
Type SYSTEMTIME
     wYear                  As Integer
     wMonth                 As Integer
     wDayOfWeek             As Integer
     wDay                   As Integer
     wHour                  As Integer
     wMinute                As Integer
     wSecond                As Integer
     wMilliseconds          As Integer
End Type
Type TIME_ZONE_INFORMATION
     Bias                   As Long       ' Basis-Zeitverschiebung in Minuten
     StandardName(1 To 64)  As Byte       ' Name der Sommerzeit-Zeitzone
     StandardDate           As SYSTEMTIME ' Beginn der Standardzeit
     StandardBias           As Long       ' Zusätzliche Zeitverschiebung der Standardzeit
     DaylightName(1 To 64)  As Byte       ' Name der Sommerzeit-Zeitzone
     DaylightDate           As SYSTEMTIME ' Beginn der Sommerzeit
     DaylightBias           As Long       ' Zusätzliche Zeitverschiebung der Sommerzeit
End Type

Declare Function GetTimeZoneInformation& Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION)


' IP-Adresse einer Internetadresse ermitteln
Public Function GetIP$(ByVal HostName$)
  Dim pHost&, HostInfo As HOSTENT
  Dim pIP&, IPArray(3) As Byte
 
  ' Informationen des Hosts ermitteln
  pHost = gethostbyname(HostName)
  If pHost = 0 Then Exit Function
 
  ' HOSTENT-Struktur kopieren
  MoveMemory HostInfo, ByVal pHost, Len(HostInfo)
 
  ' Pointer der 1. Ip-Adresse ermitteln
  ReDim IpAddress(HostInfo.hlength - 1)
  MoveMemory pIP, ByVal HostInfo.haddrlist, 4
  MoveMemory IPArray(0), ByVal pIP, 4
 
  GetIP = IPArray(0) & "." & IPArray(1) & "." & IPArray(2) & "." & IPArray(3)
End Function ' GetIP$(ByVal HostName$)

' Mit einem Server verbinden
Public Function ConnectToServer&(ByVal ServerIP$, ByVal ServerPort&)
    Dim hSock&, RetVal&, ServerAddr As SOCKADDR
    ' Socket erstellen
    hSock = socket(AF_INET, SOCK_STREAM, 0&) '  ' SOCK_DGRAM geht nicht
    If hSock = -1 Then
      ConnectToServer = -1
    Exit Function
  End If
  ' mit dem Server verbinden
  With ServerAddr
    .sin_addr = inet_addr(ServerIP)
    .sin_port = htons(ServerPort)
    .sin_family = AF_INET
  End With
  T0 = GetTickCount
  RetVal = connect(hSock, ServerAddr, Len(ServerAddr))
  If RetVal < 0 Then
    Call closesocket(hSock)
    ConnectToServer = -1
    Exit Function
  End If
  ' Rückkehren der Funktion nach dem Abfragen von ankommenden Daten erzwingen
  RetVal = ioctlsocket(hSock, FIONBIO, 1&)
  ' Socket-ID zurückgeben
  ConnectToServer = hSock
End Function ' FUNCTION ConnectToServer&(ByVal ServerIP$, ByVal ServerPort&)

' Sock/Verbindung schließen
Public Function Disconnect(ByRef Sock&)
  Call closesocket(hSock)
  Sock = 0
End Function ' Disconnect(ByRef Sock&)

' Daten senden
Public Function SendData&(ByVal Sock&, ByVal Data$)
  SendData = send(Sock, ByVal Data, Len(Data), 0&)
End Function ' SendData&(ByVal Sock&, ByVal Data$)

' Sind Daten angekommen ?
Public Function DataComeIn&(ByVal Sock&)
  Dim Tmpstr As String * 1
  DataComeIn = recv(Sock, ByVal Tmpstr, Len(Tmpstr), MSG_PEEK)
  If DataComeIn = -1 Then
    DataComeIn = WSAGetLastError()
  End If
End Function ' DataComeIn&(ByVal Sock&)

' Daten ermitteln
Public Function GetData$(ByVal Sock&)
  Dim Tmpstr As String * 4096, RetVal&
  RetVal = recv(Sock, ByVal Tmpstr, Len(Tmpstr), 0&)
  If RetVal > -1 Then
   GetData = Left$(Tmpstr, RetVal)
  End If
  T1 = GetTickCount
  TimeDelay = (T1 - T0) / 1000 / 2
End Function ' GetData$(ByVal Sock&)

Public Function SyncClock(tStr$) As Date
    Dim NTPTime#
    Dim LngTimeFrom1990&
    Dim Monat$
    Dim UTCDATE As Date
    tStr = Trim$(tStr)
    If Len(tStr) <> 4 Then
      SyncClock = CDate(0)
      Exit Function
    End If
    NTPTime = Asc(Left$(tStr, 1)) * 256 ^ 3 + _
              Asc(Mid$(tStr, 2, 1)) * 256 ^ 2 + _
              Asc(Mid$(tStr, 3, 1)) * 256 ^ 1 + _
              Asc(Right$(tStr, 1))
    LngTimeFrom1990 = NTPTime - 2840140800#
    UTCDATE = DateAdd("s", CDbl(LngTimeFrom1990 + CLng(TimeDelay)) + CurrentBias() * 60, #1/1/1990#)
    SyncClock = UTCDATE
End Function ' SyncClock(tStr$) As Date

Public Function holZeit(ByRef Server$, ByRef verzoeg!) As Date
  Dim ServerIP As String
  Const toleranz! = 0.9 ' Zahl der Sekunden, nach der der Zeitserver hintangestellt wird
  ' eventuell vorherigen Sock schließen
  If hSock < 0 Then
    Call Disconnect(hSock)
  End If
  ' ServerIP ermitteln
'  ServerIP = GetIP("www.vbapihelpline.de")
  Const servz = 16
  Dim serv$(servz)
    serv(0) = "ntp2.fau.de"
    serv(1) = "timeserver.rwth-aachen.de" ' "134.130.4.17"
    serv(2) = "time.ien.it"
    serv(3) = "hp.rz.uni-potsdam.de"
    serv(4) = "ntp3.fau.de"
    serv(5) = "ptbtime1.ptb.de"
    serv(6) = "ptbtime2.ptb.de"
    serv(7) = "ha2.hrz.uni-giessen.de"
    serv(8) = "ntp1.ptb.de" ' "192.53.103.108"
    serv(9) = "ntp1.lrz-muenchen.de" ' "129.187.254.32" 'lrz
    serv(10) = "ts2.aco.net"
    serv(11) = "nist.expertsmi.com"
    serv(12) = "wolfnisttime.com"
    serv(13) = "nist1-atl.ustiming.org"
    serv(14) = "nist1-pa.ustiming.org"
    serv(15) = "nist1-nj.ustiming.org"
    serv(16) = "time.hko.hk"
    
  Dim i%, akti%, diff$
  Dim cR As New Registry
  RegPos = RegWurzel & App.EXEName & "\Internetzeit"
  diff = cR.ReadKey("Differenz", RegPos, HKEY_LOCAL_MACHINE)
  If diff = vNS Then diff = "0"
  For i = 0 To servz
    akti = (i + CLng(diff)) Mod servz
    ServerIP = GetIP(serv(akti))
    If ServerIP <> "" Then
'  Verbinden mit dem Server
'  hSock = ConnectToServer(ServerIP, 80)
        hSock = ConnectToServer(ServerIP, 37) ' UDP: 123, geht aber nicht
        If hSock = -1 Then
'            MsgBox "Verbindung mit dem Server ist fehlgeschlagen"
            hSock = 0
        Else
            Dim ZeitS$
            Do Until ZeitS <> ""
                ZeitS$ = GetData(hSock)
                If TimeDelay > toleranz Then
                    cR.WriteKey (akti + 1) Mod servz, "Differenz", RegPos, HKEY_LOCAL_MACHINE, REG_DWORD
                    Exit Do
                End If
            Loop
        End If
    End If
    If ZeitS <> "" Then
         holZeit = SyncClock(ZeitS)
         Server = serv(akti)
         verzoeg = TimeDelay
         Exit For
    End If
 Next i
  ' Anfrage für den Abruf eines Dokuments senden
'  Call SendData(hSock, "GET http://www.vbapihelpline.de/index.php HTTP/1.1" & vbCrLf)
'  Call SendData(hSock, "Host: LonelySuicide666" & vbCrLf)
'  Call SendData(hSock, "User-Agent: LS666 HTTP-Client" & vbCrLf & vbCrLf)
  ' Empfang abfragen
'  Timer1.INTERVAL = 1
'  Timer1.Enabled = True
End Function ' holZeit(ByRef Server$, ByRef Zeit!) As Date

Function CurrentBias%()
'// Gibt die aktuelle Zeitverschiebung
'// gegenüber GMT-Uhrzeit in Minuten zurück.
    Dim udtTZI As TIME_ZONE_INFORMATION
    Dim RetVal&
    RetVal = GetTimeZoneInformation(udtTZI)
    With udtTZI
        If RetVal = TIME_ZONE_ID_DAYLIGHT Then
              CurrentBias = -(.Bias + .DaylightBias)
        Else: CurrentBias = -(.Bias + .StandardBias)
        End If
    End With
End Function ' CurrentBias%()

Public Sub trenne()
  Call Disconnect(hSock)
  Call WSACleanup
End Sub ' trenne()

Public Function starte&()
  Dim WSD As WSAData
  starte = WSAStartup(&H202, WSD)
End Function ' starte&()


Public Function InetZeit() As Date
 Dim Server$, verzoeg!
 Call starte
 InetZeit = holZeit(Server, verzoeg)
 Call trenne
End Function ' InetZeit() As Date
