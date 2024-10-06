Attribute VB_Name = "Fax"
Option Explicit
'Dim fx, fxa
Const fllNONE& = 0 'The fax server does not log events.
Const fllMIN& = 1  'The fax server logs only severe failure events, such as errors.
Const fllMED& = 2  'The fax server logs events of moderate severity, as well as severe failure events. This would include errors and warnings.
Const fllMAX& = 3  'The fax server logs all events.

Declare Function GetCurrentProcessId Lib "kernel32" () As Integer

Const fdrmNO_ANSWER = 0 'The device will not answer the call.
Const fdrmAUTO_ANSWER = 1 'The device will automatically answer the call.
Const fdrmMANUAL_ANSWER = 2 'The device will answer the call only if made to do so manually.
Public altAusgabe(0 To 3) As New CString
Public Cpt$
#Const obmitWMi = 0
#If obmitWMi Then
Public WMIreg As SWbemObjectEx ' %windir%\system32\wbem\wbemdisp.tlb
#End If
Public PatDokDirekt$
Public FaxeZwischen$
Public ErrDes$
Public ComputerName$
'Dim FSO As New FileSystemObject
'Dim fpos& ' Fehlerposition -> soll in haupt.bas o.ä. jeweils geschrieben werden

'Private Declare Function FaxGetConfiguration Lib "winfax.dll" Alias "FaxGetConfigurationA" _
 (ByVal FaxHandle As Long, ByRef pFaxConfiguration As FAXCONFIGURATION) _
 As Long

'Private Type FAX_TIME
'    Hour As Long
'    Minute As Long
'End Type
'Private Type FAXCONFIGURATION
'    SizeOfStruct As Long
'    Retries As Long
'    RetryDelay As Long
'    Branding As Long
'    DirtyDays As Long
'    UseDeviceTsid As Long
'    ServerCp As Long
'    PauseServerQueue As Long
'    StartCheapTime As FAX_TIME
'    StopCheapTime As FAX_TIME
'    ArchiveOutgoingFaxes As Long
'    ArchiveDirectory As String
'    InboundProfile As String
'End Type
'Private hFax As Long

Function GetPCName$()
 If LenB(ComputerName) = 0 Then
  ComputerName = CptName
 End If
 GetPCName = ComputerName
End Function ' GetPCName()
Function FxAutoSet() ' Ist das gemeinsam?
 Dim fxs As FAXCOMEXLib.FaxServer, Verz$
 On Error GoTo fehler
 Set fxs = New FAXCOMEXLib.FaxServer
 FxConnect fxs, GetPCName
 Verz = fxs.Folders.IncomingArchive.ArchiveFolder
 Verz = Left$(Verz, Len(Verz) - 12) ' minus \MSFax\Inbox
 Call fxset(Verz)
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxautoset/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FxAutoSet
'Public Function Ende()
''  oEnvSystem.Environment("NVerb") = vns
'  End
'End Function

Function fxset(Optional Verz$)
 Dim fx As FAXCOMEXLib.FaxServer      'CreateObject("FAXCOMEX.FaxServer")
 Dim fxs As faxcomlib.FaxServer       'CreateObject("faxserver.faxserver")
 Dim fxc As FAXCONTROLLib.FaxControl  'CreateObject("faxcontrol.faxcontrol")
 Dim fxps As faxcomlib.FaxPorts
 Dim fxp As faxcomlib.FaxPort
 Dim fxst As faxcomlib.FaxStatus
 Dim i%, j%
 Dim breg As New Registry
 Dim ByteArray(1) As Byte, T1 As Date, T2 As Date
 Dim BAVar() As Byte
' ByteArray = StrConv("Hallo, nanu, hört es schon auf??", vbFromUnicode) '...c.Value = ByteArray
 Static fehlg1%, fehlg2%
 Dim modul$
 Dim FaxFehler%
 
 On Error GoTo fehler
 ' Nur ein Programm gleichzeitig soll new faxcontrollib.faxcontrol aufrufen => Aktives Programm schreibt sich in die Registry
 breg.ClassKey = HKEY_LOCAL_MACHINE
 breg.SectionKey = "SOFTWARE\GSProducts\fxset"
 breg.ValueType = REG_BINARY
 T1 = Now
 Do
  breg.ReadKey
  If LenB(breg.Value) = 0 Then Exit Do
  If breg.Value(0) <> 1 Then Exit Do
  T2 = Now
  If (T2 - T1) * 24 * 60 > 0.1 Then Exit Do ' nach 5 min darf er's mal versuchen
  Sleep 1000 ' damit er nicht dauernd arbeitet
 Loop
 ByteArray(0) = 1
 breg.Value = ByteArray
 breg.CreateKey
 
vorfax:
 Set fxc = New FAXCONTROLLib.FaxControl
 If Not fxc.IsFaxServiceInstalled Then fxc.InstallFaxService
 FaxFehler = -1
 If Not fxc.IsLocalFaxPrinterInstalled Then fxc.InstallLocalFaxPrinter
 FaxFehler = 0
 ByteArray(0) = 0
 breg.WriteKey ByteArray, , "SOFTWARE\GSProducts\fxset", HKEY_LOCAL_MACHINE, REG_BINARY
 
 FPos = 1
 Set fxs = Nothing
 On Error GoTo modul
 Set fxs = New faxcomlib.FaxServer
 On Error GoTo fehler
 If fxs Is Nothing Then
  modul = "faxcomlib"
  Set fxs = New faxcomlib.FaxServer
  modul = ""
 End If
 FPos = 2
 If Not fxs Is Nothing Then
 With fxs
  If LenB(Verz) <> 0 Then .ArchiveDirectory = Verz + "\MSFax\Inbox"
  .ArchiveOutboundFaxes = -1
  .Branding = -1 ' Banner
  .DirtyDays = 3 ' wird nicht da angezeigt wo vermutet, lässt sich nicht ändern
  .Retries = 3
  .RetryDelay = 5
  .ServerCoverpage = 0
  .UseDeviceTsid = -1
 End With
 FPos = 3
 If FxConnect(fxs, vNS) Then Exit Function
 FPos = 4
 Set fxps = fxs.GetPorts
 FPos = 5
 Dim faxnr%
 faxnr = 1
 With fxps
  FPos = 6
  For i = 1 To .COUNT
   If InStrB(fxps.Item(i).name, "56") <> 0 Then faxnr = i ' MicroLink 56k PCI
  Next i
' For i = 1 To .Count
   FPos = 7
   On Error Resume Next ' 29.6.15
   Set fxp = fxps.Item(faxnr)
   On Error GoTo fehler
   FPos = 8
   If fxp Is Nothing Then Exit Function ' kein Modem
   With fxp
    FPos = 9
'   ' debug.print.CanModify
    .Csid = "DiabDach 08131616381"
'   ' debug.print.Name
    ' debug.print.Priority
    .Rings = 1 ' steht bei "Geräte" sowohl innen als auch außen, zulässiger Bereich = 1-99
    On Error Resume Next
    Err.Clear
    .Receive = 1 ' dann geht der Haken von "Manuell" auf "Automatisch"
    If Err.Number <> 0 Then
     If Not fehlg1 Then
'      Call Shell(App.Path & "\nachricht.exe " & "Achtung: Fax '" & fxp.name & "' ließ sich nicht auf Empfangen schalten.")
'      SuSh App.Path & "\nachricht.exe " & "Achtung: Fax '" & fxp.Name & "' ließ sich nicht auf Empfangen schalten.", 0, , 0, 1
      rufauf App.Path & "\nachricht.exe", "Achtung: Fax '" & fxp.name & "' ließ sich nicht auf Empfangen schalten.", , , 0, 1
      fehlg1 = True
     End If
     Err.Clear
    End If
    .Send = -1
    If Err.Number <> 0 Then
     If Not fehlg2 Then
'      Call Shell(App.Path & "\nachricht.exe " & "Achtung: Fax '" & fxp.name & "' ließ sich nicht auf Senden schalten.")
'      SuSh App.Path & "\nachricht.exe " & "Achtung: Fax '" & fxp.Name & "' ließ sich nicht auf Senden schalten.", 0, , 0, 1
      rufauf App.Path & "\nachricht.exe", "Achtung: Fax '" & fxp.name & "' ließ sich nicht auf Senden schalten.", , , 0, 1
      fehlg2 = True
     End If
     Err.Clear
    End If
    On Error GoTo fehler
    .Tsid = "DiabDach 08131616381"
    FPos = 10
    On Error Resume Next
    Set fxst = fxp.GetStatus ' wohl aktuell gesendetes Fax
    On Error GoTo fehler
    FPos = 11
    Dim rms As faxcomlib.FaxRoutingMethods
    FPos = 12
    Dim rm As faxcomlib.FaxRoutingMethod
    Set rms = fxp.GetRoutingMethods
    FPos = 13
    For j = 1 To rms.COUNT
     FPos = 14
     Set rm = rms.Item(j)
      FPos = 15
     ' debug.printrm.DeviceId
     ' debug.printrm.DeviceName
     ' debug.printrm.Enable
     ' debug.printrm.ExtensionName
     ' debug.printrm.FriendlyName
     ' debug.printrm.FunctionName
     ' debug.printrm.Guid
     ' debug.printrm.ImageName
     ' debug.printrm.RoutingData ' hier schreibgeschützt
    Next j
   End With
'  Next i
 End With
 
 FPos = 16
 Set fx = New FAXCOMEXLib.FaxServer
 FPos = 17
 With fx
  FxConnect fx, Cpt
  FPos = 18
  With .Folders
   FPos = 19
   With .OutgoingArchive
    FPos = 20
'    .ArchiveFolder = Verz + "\MSFax\SentItems"
    .UseArchive = True
    .AgeLimit = 0 ' scheint nicht angezeigt zu werden
    .SizeQuotaWarning = False
    .HighQuotaWaterMark = -1
    .LowQuotaWaterMark = -1
    .Save
   End With
   On Error Resume Next
    With .OutgoingArchive
     If LenB(Verz) <> 0 Then .ArchiveFolder = Verz & "\MSFax\SentItems"
     FPos = 21
     .Save
    End With
   On Error GoTo fehler
   FPos = 22
   With .IncomingArchive
'    .ArchiveFolder = Verz + "\MSFax\Inbox"
    .UseArchive = True
    .AgeLimit = 0 ' scheint nicht angezeigt zu werden
    .SizeQuotaWarning = False
    .Save
   End With
   FPos = 23
   On Error Resume Next
   With .IncomingArchive
    If LenB(Verz) <> 0 Then .ArchiveFolder = Verz & "\MSFax\Inbox"
    .Save
   End With
   FPos = 24
   On Error GoTo fehler
   With .IncomingQueue
    .Blocked = False
    .Save
   End With
   FPos = 25
   With .OutgoingQueue
    .Blocked = False ' scheint nicht angezeigt zu werden
    .AgeLimit = 0 ' bei Geräte unter Bereinigen, wenn <> 0, dann angekreuzt
    .Save
   End With
  End With ' folders
  FPos = 26
  With .LoggingOptions
   With .EventLogging
    .InitEventsLevel = fllMAX
    .InboundEventsLevel = fllMAX
    .OutboundEventsLevel = fllMAX
    .GeneralEventsLevel = fllMAX
    .Save
   End With
   FPos = 27
   With .ActivityLogging
    .LogIncoming = fllMAX
    .LogOutgoing = fllMAX
    .DatabasePath = Environ("allusersprofile") + "\Anwendungsdaten\Microsoft\Windows NT\MSFax\ActivityLog"
    .Save
   End With
  End With
  FPos = 28
  Dim fdev As FAXCOMEXLib.FaxDevices
NumberOfActiveFaxDevices:
  Set fdev = fx.GetDevices
  FPos = 29
  With fdev
   For j = 1 To .COUNT + 1 ' bei Windows XP Home geht nur ein Gerät aktiv
    If j = .COUNT + 1 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
     MsgBox "Fehler in '" & App.Path & "/fxset' beim Speichern eines 'Faxcomexlib.faxdevices', .count: " & .COUNT & vbCrLf & "Höre auf!"
     ProgEnde
    End If
    With .Item(j)
     .RingsBeforeAnswer = 1 ' ist das selbe wie oben
     .Csid = "DiabDach 08131616381"
     .Tsid = "DiabDach 08131616381"
     .SendEnabled = -1
     .ReceiveMode = fdrmAUTO_ANSWER
     .Description = Left$(.DeviceName, InStr(.DeviceName, Space(1))) & " Fax von " & GetPCName & " für MSFax"
     On Error Resume Next
     Err.Clear
     .Save
     If Err.Number = 0 Then Exit For
     On Error GoTo fehler
    End With
   Next j
  End With
  FPos = 30
  Dim fibrms As FAXCOMEXLib.FaxInboundRoutingMethods
  Dim fibrm As FAXCOMEXLib.FaxInboundRoutingMethod
  Set fibrms = fx.InboundRouting.GetMethods
  FPos = 31
  For i = 0 To fibrms.COUNT
   With fibrms.Item(i)
    ' debug.print.ExtensionFriendlyName
    ' debug.print.ExtensionImageName
    ' debug.print.FunctionName
    ' debug.print.Guid
    ' debug.print.Name
    ' debug.print.Priority
   End With
  Next i
  FPos = 32
  Dim fibres As FAXCOMEXLib.FaxInboundRoutingExtensions
  Dim fibre As FAXCOMEXLib.FaxInboundRoutingExtension
  Set fibres = fx.InboundRouting.GetExtensions
  For i = 0 To fibres.COUNT
   Set fibre = fibres(i)
   With fibre
     ' debug.print.Debug
     ' debug.print.FriendlyName
     ' debug.print.ImageName
     ' debug.print.InitErrorCode
     ' debug.print.MajorBuild
     ' debug.print.MajorVersion
     ' debug.print.Methods(0)
     ' debug.print.Methods(1)
     ' debug.print.MinorBuild
     ' debug.print.MinorVersion
     ' debug.print.Status
     ' debug.print.UniqueName
   End With
  Next i
  
 End With
 FPos = 33
 
 ' Kopie der Faxe im Patientenordner speichern
 'Dim wmireg As SWbemObjectEx
#If obmitWMi Then
  Dim Result&, arra
  arra = Arr(PatDokDirekt) ' c:\P
  If WMIreg Is Nothing Then Set WMIreg = GetObject("winmgmts:root\default:StdRegProv")
 ' Wert eintragen
  Result = WMIreg.setbinaryvalue(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{92041a90-9af2-11d0-abf7-00c04fd91a4e}", arra)
 ' aktivieren
  Result = WMIreg.setbinaryvalue(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{aacc65ec-0091-40d6-a6f3-a2ed6057e1fa}", Array(2, 0, 0, 0))
#Else
'  Dim breg As New Registry
  breg.ClassKey = HKEY_LOCAL_MACHINE
  breg.SectionKey = "SOFTWARE\Microsoft\Fax\TAPIDevices"
'  breg.ValueType = REG_BINARY
'  breg.ValueKey = "{92041a90-9af2-11d0-abf7-00c04fd91a4e}"
''  BAVar = StrConv(FaxeZwischen, vbFromUnicode) '...c.Value = ByteArray
'  breg.Value = FaxeZwischen
  Dim ens$(), ec&
  breg.EnumerateSections ens, ec
  For i = 0 To ec
   Select Case ens(i)
'    Case "014BFAB1", "00bd262e"
    Case Else
     breg.WriteKey FaxeZwischen, "{92041a90-9af2-11d0-abf7-00c04fd91a4e}", "SOFTWARE\Microsoft\Fax\TAPIDevices\" & ens(i), HKEY_LOCAL_MACHINE, REG_BINARY
     breg.WriteKey Array(2, 0, 0, 0), "{aacc65ec-0091-40d6-a6f3-a2ed6057e1fa}", "SOFTWARE\Microsoft\Fax\TAPIDevices\" & ens(i), HKEY_LOCAL_MACHINE, REG_BINARY
    End Select
  Next i
'  Call fBiSpei(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{92041a90-9af2-11d0-abf7-00c04fd91a4e}", BAVar)
'  Call fBiSpei(HLM, "SOFTWARE\Microsoft\Fax\TAPIDevices\014BFAB1", "{aacc65ec-0091-40d6-a6f3-a2ed6057e1fa}", &O1)
#End If
 FPos = 34
End If ' not fxs is nothing
' Call fStSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Folder", EigDatDirekt + "\MSFax\Inbox")
 Call fDWSpei(HLM, "SOFTWARE\Microsoft\Fax\Inbox", "Use", 1)
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Dr.Th.Kothny + G.Schade")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Address", "Mittermayerstraße 13" + vbCrLf + "85221 Dachau")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "City", "Dachau")
' Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Schade")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Company", "Praxis")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "Country", "Deutschland")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FaxNumber", "08131 616381")
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "HomePhone", "08131 616380")
 FPos = 35
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "mailbox", "diabetologie@dachau-mail.de")
 Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "MonitorOnReceive", 1)
 Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "MonitorOnSend", 1)
 Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyIncomingCompletion", 1)
 Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyOutgoingCompletion", 1)
 Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "NotifyProgress", 1)
 Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "ZIP", "85221")
' Call fStSpei(HCU, "SOFTWARE\Microsoft\Fax\UserInfo", "FullName", "Gerald Schade")
' Call fDWSpei(HCU, "SOFTWARE\Microsoft\Fax", "Dirty Days", 0) ' zu ändern über outgoing folder
' Call fStSpei(HCU, "SYSTEM\CurrentControlSet\Control\Print\Printers\Fax", "Location", cpt) ' zu ändern was weiß ich wo
 Exit Function
modul:
Dim dll$
If WV < win_vista Then
 FileCopy "u:\programmierung\fxscom.dll", Environ("windir") + "\system32"
 dll = Environ("windir") & "\system32\fxscom.dll"
' Shell ("regsvr32.exe " & dll)
' SuSh "regsvr32.exe " & dll, 2
  rufauf "regsvr32.exe", dll, 2, , , 0
Else
 KopDat "u:\programmierung\fxscom.dll", Environ("windir") + "\sysWOW64"
 dll = Environ("windir") & "\syswow64\fxscom.dll"
' ShellaW doalsad & acceu & AdminGes & " cmd /e:on /c regsvr32 " & Chr$(34) & dll & Chr$(34), vbHide, , 10000
' SuSh "cmd /e:on /c regsvr32 " & Chr$(34) & dll & Chr$(34), 2
 rufauf "cmd", "/e:on /c regsvr32""" & dll & """", 2, , , 0
End If
schließ_direkt ("regsvr32")
Resume
fehler:
Dim ErrNr&
ErrNr = Err.Number
If FPos = 29 And ErrNr = -2147214494 Then
 Resume NumberOfActiveFaxDevices
ElseIf ErrNr = -2147467259 Then 'Automatisierungsfehler
 Call faxRestart(fxs, fxc, Verz)
 Resume
Else
 If FaxFehler Then Exit Function
 Select Case MsgBox("FNr: " & FNr & IIf(FPos = 29 And ErrNr = -2147214494, ", j = " & j, vNS) & ", ErrNr: " & CStr(ErrNr) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + "( Modul: " + modul + ") " + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxset/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal")
   Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End If
End Function ' fxset

Function faxRestart(Optional fxs As faxcomlib.FaxServer, Optional fxc As FAXCONTROLLib.FaxControl, Optional Verz$)
   On Error Resume Next
   fxs.Disconnect
   If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo fehler
#If obmitWMi Then
    Dim objWMI As SWbemServicesEx, objDienst As SWbemObjectEx, rv&
    On Error GoTo fehler
    Set objWMI = GetObject("winmgmts:\\" & GetPCName & "\ROOT\CIMV2")
    Set objDienst = objWMI.Get("Win32_Service.name='fax'")
    With objDienst
     If .AcceptStop Then
      rv = objDienst.StopService ' Dienst stoppen bzw. beenden
     Else
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
      MsgBox "Fax lässt sich nicht stoppen." & vbCrLf & "Der Dienst läuft möglicherweise nicht."
     End If
    End With
#Else
' hier muß noch neue Dienstklasse eingebaut werden
'   Dim erg&
'   erg = ShellaW("net stop fax", vbHide, , 100000)
'   erg = SuSh("net stop fax", 2)
    rufauf "cmd", "/c net stop fax", , , 0, 0
    rufauf "cmd", "/c net stop fax", 2, , 0, 0
'   Call WarteAufNicht("fxssvc.exe", 1000)
#End If
   End If ' err.number <> 0
   Set fxc = New FAXCONTROLLib.FaxControl
   fxc.InstallFaxService
   fxc.InstallLocalFaxPrinter
   Set fxs = New faxcomlib.FaxServer
   With fxs
    If LenB(Verz) <> 0 Then .ArchiveDirectory = Verz & "\MSFax\Inbox"
    .ArchiveOutboundFaxes = -1
    .Branding = -1 ' Banner
    .DirtyDays = 3 ' wird nicht da angezeigt wo vermutet, lässt sich nicht ändern
    .Retries = 3
    .RetryDelay = 5
    .ServerCoverpage = 0
    .UseDeviceTsid = -1
   End With
   FxConnect fxs, vNS
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in faxRestart/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' faxRestart

Function ReStart()
 On Error GoTo fehler
   Call GetProcessCollection(2, "fxssvc.exe")
   Call WarteAufNicht("fxssvc.exe", 10)
'   Shell "fxssvc.exe"
'   Call WarteAuf("fxssvc", 10)
'   SuSh "fxssvc.exe", 2, , 0
   rufauf "fxsvc.exe", , 2, , 0, 0
   rufauf "fxsvc.exe", 2, 2, , 0, 0
   Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ReStart/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ReStart

Function FxConnect(fxs, Cpt$)
Dim ErrZahl&
 On Error GoTo fehler
' es folgt die einzige Stelle innerhalb der verwendenden Programme AutoFax, FaxAkt und FaxDopp,
' an der .connect aufgerufen wird; Typ von fxs verschieden
 FxConnect = 0
 fxs.Connect Cpt
 Exit Function
fehler:
Select Case ErrZahl
 Case Is < 10
  If Err.Number = -2147023174 Or Err.Number = 462 Then ' Connection to fax server failed.
   Call ReStart
  Else
   meld (App.Path & ": Fehler bei fxs.Connect:" & vbCrLf & "Err.Number: " & Err.Number & vbCrLf & "Err.Description: " & Err.Description)
   Call ReStart
  End If
  ErrZahl = ErrZahl + 1
  Resume
 Case Else
' Call Shell(App.Path & "\nachricht.exe " & "fxs.connect " & Cpt & ": Fehler " & CStr(Err.Number) & vbCrLf & Err.Description & vbCrLf & "Arbeite ohne Fax")
' SuSh App.Path & "\nachricht.exe " & "fxs.connect " & Cpt & ": Fehler " & CStr(Err.Number) & vbCrLf & Err.Description & vbCrLf & "Arbeite ohne Fax", 0, , 0, 1
 rufauf App.Path & "\nachricht.exe", "fxs.connect " & Cpt & ": Fehler " & CStr(Err.Number) & vbCrLf & Err.Description & vbCrLf & "Arbeite ohne Fax", , , 0, 1
 FxConnect = 1
 Resume Next
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FxConnect/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Select
End Function ' FxConnect

Function Arr(Pfad$) ' Baut nach jedem Buchstaben eine Lücke ein, so wie es offenbar die Routing-Funktion will
 Dim i%, j%, arra%()
 On Error GoTo fehler
 ReDim arra(2 * Len(Pfad))
 j = 0
 For i = 0 To Len(Pfad) - 1
  arra(j) = Asc(Mid(Pfad, i + 1, 1))
  arra(j + 1) = 0
  j = j + 2
 Next i
 Arr = arra
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Arr/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Arr

Function PDFDruckerName$(Optional obPDFCreator%)
'Dim objWMIService As Object, colItems, msg$, objItem
'Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
'Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer", , 48)
'   msg = "Installierte Drucker:" & vbCr
'For Each objItem In colItems
'    msg = msg & objItem.Caption & vbCr
'Next
Dim myptr As Printer
Static PDFDrucker$
On Error GoTo fehler
If LenB(PDFDrucker) <> 0 Then
 PDFDruckerName = PDFDrucker
 Exit Function
End If
For Each myptr In Printers
 If (obPDFCreator And InStr(UCase$(myptr.DeviceName), "PDFCREATOR ZUFAXEN") <> 0) Or (Not obPDFCreator And InStr(UCase$(myptr.DeviceName), "FREEPDF") = 1 And InStrB(UCase$(myptr.DeviceName), "AUTO")) <> 0 Then
   PDFDruckerName = myptr.DeviceName
   PDFDrucker = PDFDruckerName
   Exit Function
 End If
Next
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxp/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PDFDruckerName

Function FaxSend&(RecName$, RecNum$, DateiName$, Optional Adr$, Optional Cit$, Optional Com$, Optional Coun$, Optional Dep$, Optional Homeph$, Optional obstreng%, Optional pdf$)
Dim FaxServer As New faxcomlib.FaxServer
Dim FaxDoc As New faxcomlib.FaxDoc
Dim FaxTiff As New faxcomlib.FaxTiff
Dim strFaxJob As faxcomlib.FaxJobs
Dim strFaxStatus As faxcomlib.FaxJob
Dim strFaxTiff As faxcomlib.FaxTiff
On Error GoTo fehler
 Call SetProgV
' DateiName = uverz & "test1.txt"
 Err.Clear
 FxConnect FaxServer, GetPCName
Dim i&, endung$, lenge&
lenge = Len(DateiName)
Dim Fil As File
Err.Clear
Set Fil = FSO.GetFile(DateiName)
If Err.Number <> 0 Then Exit Function
 
For i = lenge To 1 Step -1
 If Mid(DateiName, i, 1) = "." Then
  endung = LCase$(Right$(DateiName, lenge - i))
  Exit For
 End If
Next i
Dim PdfName$
PdfName = Fil.ParentFolder.Path & "\*" & Left$(Fil.name, Len(Fil.name) - Len(endung)) & "pdf"

If InStrB(DateiName, "AutofaxFehler") <> 0 Then
 Print #389, DateiName & ": schon gescheitert"
 Exit Function
End If
Select Case endung
 Case "aif", "ani", "au", "avi", "awd", "b3d", "bmp", "cam", "clp", "cpt", "crw", "cr2", "cur", "dat", "dcm", "acr", "dcx", "dds", _
 "dib", "djvu", "dll", "exe", "dxf", "dwg", "hpgl", "cgm", "svg", "ecw", "emf", "eps", "ps", "exr", "fits", "fli", "flc", "fpx", _
 "fsh", "g3", "gif", "hdp", "wdp", "hdr", "icl", "ico", "iff", "lbm", "img", "jp2", "jpg", "jpeg", "jpm", "kdc", "ldf", "lwf", _
 "mac pict", "mag", "med", "mid", "mng", "jng", "mov", "mp3", "mpg", "mpeg", "mrsid", "nlm", "nol", "ngg", "nsl", "otb", "gsm", _
 "ogg", "pbm", "pcx", "pgm", "pic", "photocd", "png", "ppm", "psd", "psp", "pvr", "qtif", "ras", "sun", "realaudio", "raw", _
 "rle", "rmi", "sff", "sfw", "sgi", "rgb", "mrsid", "swf", "flv", "snd", "tga", "tif", "tiff", "ttf", "txt", "wad", "wal", _
 "wav", "wbmp", "wmf", "wma", "wmv", "wsq", "xbm", "xpm", "dng", "eef", "nef", "orf", "raf", "mrw", "dcr", "pef", "srf", "x3f"
 If LenB(PDFDruckerName) = 0 Then
'  Call Shell(App.Path & "\nachricht.exe " & "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da Drucker 'FreePDF - auto' nicht gefunden!")
'  SuSh App.Path & "\nachricht.exe " & "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da Drucker 'FreePDF - auto' nicht gefunden!", 0, , 0, 1
  rufauf App.Path & "\nachricht.exe", "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da Drucker 'FreePDF - auto' nicht gefunden!", , , 0, 1
  Exit Function
 End If
 Debug.Print ProgVerz & "\irfanview\i_view32.exe """ & DateiName & """ /print=""" & PDFDruckerName & """"
' Shell (ProgVerz & "\irfanview\i_view32.exe """ & DateiName & """ /print=""" & PDFDruckerName & """")
' SuSh ProgVerz & "\irfanview\i_view32.exe """ & DateiName & """ /print=""" & PDFDruckerName & """"
  rufauf ProgVerz & "\irvanview\i_view32.exe""""", DateiName & """ /print=""" & PDFDruckerName & """", , , , 0
 Dim T1#, T2#
 T1 = Timer
 Do
  DoEvents
  T2 = Timer
'  pdf = dir(PdfName)
  If FileExists(PdfName) Then
    pdf = Mid$(PdfName, InStrRev(PdfName, "\") + 1)
    Exit Do
  End If
  If T2 - T1 > 120 Then
   pdf = ""
   Exit Do
  End If
 Loop
 If Not FileExists(PdfName) Then
'  Call Shell(App.Path & "\nachricht.exe " & "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da nach PDF-Erstellung mit Drucker '" & PDFDruckerName & "'" & vbCrLf & "keine Datei wie: '" & PdfName & "' gefunden!")
'  SuSh App.Path & "\nachricht.exe " & "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da nach PDF-Erstellung mit Drucker '" & PDFDruckerName & "'" & vbCrLf & "keine Datei wie: '" & PdfName & "' gefunden!", 0, , 0, 1
  rufauf App.Path & "\nachricht.exe", "Achtung: Datei '" & DateiName & "'" & vbCrLf & " nicht faxbar, da nach PDF-Erstellung mit Drucker '" & PDFDruckerName & "'" & vbCrLf & "keine Datei wie: '" & PdfName & "' gefunden!", , , 0, 1
  Exit Function
 End If
 Dim neupdf$
 If FileExists(PdfName) Then
  On Error Resume Next
  T1 = Timer
  Do ' erfüllt 2 Funktionen: erstens die PDF umbenennen, zweitens warten, bis sie wirklich fertig ist
   DoEvents
   neupdf = Left$(Fil.name, Len(Fil.name) - Len(endung)) & "pdf"
   Name Fil.ParentFolder.Path & "\" & pdf As Fil.ParentFolder.Path & "\" & neupdf
   T2 = Timer
   If FileExists(Fil.ParentFolder.Path & "\" & neupdf) Then
    pdf = neupdf
    Exit Do
   End If
   If T2 - T1 > 300 Then Exit Do
  Loop
  On Error GoTo fehler
 End If
 Case vNS
  Exit Function
End Select
Debug.Print "pdf: " & pdf, True
Dim dname$

If LenB(pdf$) = 0 Then
 dname = DateiName
Else
 dname = Fil.ParentFolder.Path & "\" & pdf
End If
nochmal:
Set FaxDoc = FaxServer.CreateDocument(dname)
    FaxDoc.BillingCode = "Rechnungsnummer 381"
    FaxDoc.CoverpageName = vNS
    On Error Resume Next
    FaxDoc.CoverpageNote = FSO.GetFile(dname).name
    If Err.Number = 53 Then Exit Function Else If Err.Number <> 0 Then FaxDoc.CoverpageNote = Fil.name ' Datei nicht gefunden, 12.9.12
    On Error GoTo fehler
    FaxDoc.CoverpageSubject = FaxDoc.CoverpageNote
    FaxDoc.DiscountSend = 0
    FaxDoc.DisplayName = FaxDoc.CoverpageNote  ' "G.Schade"
    FaxDoc.EmailAddress = "diabetologie@dachau-mail.de"
    FaxDoc.FaxNumber = RecNum
    FaxDoc.RecipientAddress = vNS
    FaxDoc.RecipientCity = vNS
    FaxDoc.RecipientCompany = "Praxis"
    FaxDoc.RecipientCountry = "D"
    FaxDoc.RecipientDepartment = vNS
    FaxDoc.RecipientHomePhone = vNS
    FaxDoc.RecipientName = RecName
    FaxDoc.RecipientOffice = vNS
    FaxDoc.RecipientOfficePhone = vNS
    FaxDoc.RecipientState = "Bayern"
    FaxDoc.RecipientTitle = vNS
    FaxDoc.RecipientZip = vNS
    FaxDoc.SendCoverpage = 0
    FaxDoc.SenderAddress = "Mittermayerstraße 13"
    FaxDoc.SenderCompany = "Diabetologische Gemeinschaftspraxis"
    FaxDoc.SenderDepartment = "Schreibbüro"
    FaxDoc.SenderFax = "08131 616381"
    FaxDoc.SenderHomePhone = "616380"
    FaxDoc.SenderName = "Dr.Th.Kothny + G.Schade"
    FaxDoc.SenderOffice = "Praxis"
    FaxDoc.SenderOfficePhone = "616380"
    FaxDoc.SenderTitle = vNS
    FaxDoc.ServerCoverpage = 1
    FaxDoc.FileName = dname
    If obstreng = 0 Then On Error Resume Next
    Err.Clear
    FaxSend = FaxDoc.Send
    If Err.Number <> 0 Then
     ErrDes = Err.Description ' immer: "Die Methode 'Send' für das Objekt 'IFaxDoc' ist fehlgeschlagen"
     Print #389, dname & ": Fehler!"
     Dim Ziel$
     Ziel = Left(DateiName, Len(DateiName) - Len(endung) - 1) & " AutofaxFehler." & endung ' & "_" & Replace$(Replace$(Replace$(ErrDes, Chr$(32), "_"), "\", "-"), ":", ";") & "." & endung
     Call schließ_direkt("Adobe Acrobat - [1")
     Name DateiName As Ziel ' geht aus mir unbekannten Gründen nicht
    Else
'     Print #389, dname & ": Erfolg!"
     Call schließ_direkt("Adobe Acrobat - [1")
    End If
    On Error GoTo fehler
    DoEvents
    
'    MsgBox FaxServer.ArchiveDirectory
  
Set strFaxJob = FaxServer.GetJobs()
DoEvents
Set strFaxStatus = strFaxJob.Item(1)
    
On Error Resume Next

Set FaxServer = Nothing
Set FaxDoc = Nothing
Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FaxSend

#If einstmals Then
Function fxpfu()
Dim myptr As Printer
On Error GoTo fehler
For Each myptr In Printers
   If myptr.DeviceName = "Fax" Then
      Set Printer = myptr
      Exit For
   End If
Next
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fxp/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fxp

Public Function testAktiv%(Optional FI As Form)
 Dim col As New Collection, AppZahl%, i%, Cpid&
 On Error GoTo fehler
 ' fi.Ausgeb "vor GetProcessCollection", 3, 0
 Set col = GetProcessCollection
 ' fi.Ausgeb "nach GetProcessCollection", 3, 0
 Cpid = GetCurrentProcessId
 For i = 1 To col.COUNT
  ' fi.Ausgeb "testAktiv: " & col(i) & " " & App.EXEName, 3, 0
  If InStrB(col(i), "|" & Cpid) = 0 Then ' nicht aktuelle Anwendung
   If InStrB(col(i), App.EXEName) <> 0 Then
    ' fi.Ausgeb "InstrB(" & col(i) & "," & "|" & Cpid & ")=0", 3, 0
    ' fi.Ausgeb "testAktiv: " & col(i) & ": beende, da Programm schon läuft", 3, 0
    AppZahl = AppZahl + 1
    If AppZahl > 0 Then
'   MsgBox "Programm wird bereits ausgeführt!"
     testAktiv = True
     Exit For
    End If
   ElseIf InStrB(UCase$(col(i)), "NVINI") <> 0 Or InStrB(UCase$(col(i)), "NVERB") <> 0 Then
    ' fi.Ausgeb "testAktiv: " & col(i) & ": beende, da NVIni läuft", 3, 0
    testAktiv = True
    Exit For
   End If
  End If
 Next i
 ' fi.Ausgeb "am Ende von testAktiv", 3, 0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in testAktiv/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testAktiv
#End If

