Attribute VB_Name = "ComputerTools"
' erfordert GetProcColl.bas
' erfordert ein Formular fürIcon
Option Explicit
Declare Function environonmentVariable& Lib "kernel32.dll" Alias "environonmentVariableA" (ByVal lpName$, ByVal lpBuffer$, ByVal nSize&)
Declare Function SetEnvironmentVariable& Lib "kernel32.dll" Alias "SetEnvironmentVariableA" (ByVal lpName$, ByVal lpValue$)

Declare Function gethostbyname& Lib "wsock32.dll" (ByVal HostName$)
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Declare Function WSACleanup& Lib "wsock32.dll" ()

Public WV As WindowsVersion

Const WSADescription_Len As Long = 256
'Const SOCKET_ERROR  AS Long = -1
Const WSASYS_Status_Len As Long = 128
Type WinSocketDataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Declare Function WSAStartup& Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAData As WinSocketDataType)
Declare Function Beep& Lib "kernel32" (ByVal dwFreq&, ByVal dwDuration&)
Declare Function MessageBeep Lib "user32.dll" (ByVal wType As Long) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (lpszName As Any, ByVal hModule&, ByVal dwFlags&) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)


'Dim fpos& ' Fehlerposition -> soll in haupt.bas o.ä. jeweils geschrieben werden
Type HostDeType
    hname As Long
    haliases As Long
    haddrtype As Integer
    hlength As Integer
    haddrlist As Long
End Type

Const WS_VERSION_REQD As Long = &H101&

#If vorWin7 Then
Public Enum WindowsVersion
       WIN_OLD
       WIN_95
       WIN_98
       WIN_ME
       WIN_NT_3x
       win_nt_4x
       win_2k
       win_xp
       win_xP_home
       WIN_2003
       win_vista
       win7
       win8
       win9
       win10
End Enum
Private Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128 ' Service Pack
End Type
Private Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128 ' Service Pack
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer
        wProductType As Byte
        wReserved As Byte
End Type

Declare Function GetVersionEx1& Lib "kernel32.dll" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO)
Declare Function GetVersionEx2& Lib "kernel32.dll" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFOEX)
#Else
Public Enum WindowsVersion
       win_vista = 6
End Enum
#End If

'Declare FUNCTION IsWindowsXPOrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsXPSP1OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsXPSP2OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsXPSP3OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsVistaOrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsVistaSP1OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsVistaSP2OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindows7OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindows7SP1OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindows8OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindows8Point1OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindows10OrGreater Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsServer Lib "ntdll.dll" () AS Boolean
'Declare FUNCTION IsWindowsVersionOrGreater Lib "ntdll.dll" (ByVal wMajorVersion&, ByVal wMinorVersion&, ByVal wServicePackMajor&) AS Boolean

'' os product type VALUES
'Public Const VER_NT_WORKSTATION AS Long = &H1 '&gt;&gt;&gt;!!! NOT THE SERVER &lt;&lt;&lt;
'Public Const VER_NT_DOMAIN_CONTROLLER AS Long = &H2
'Public Const VER_NT_SERVER AS Long = &H3
'
'Public Type OSVERSIONINFOEX1
'dwOSVersionInfoSize AS Long
'dwMajorVersion AS Long
'dwMinorVersion AS Long
'dwBuildNumber AS Long
'dwPlatformId AS Long
'szCSDVersion AS String * 128
'wServicePackMajor As Integer
'wServicePackMinor As Integer
'wSuiteMask As Integer
'wProductType AS Byte
'wReserved AS Byte
'End Type

'Public Declare FUNCTION GetVersionEx3 Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation AS Any) AS Long


Public FZ%     ' Freigabezahl
Public FNam$() ' FreigabeNamen
Public FInh$() ' FreigabeInhalt

Private Type SHITEMID
  cb   As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation& Lib "Shell32" (ByVal hwndOwner&, ByVal nFolder&, ByRef ppidl As ITEMIDLIST)
Private Declare Function SHGetPathFromIDList& Lib "Shell32" (ByVal pidList&, ByVal lpBuffer$)

Private Const S_OK = 0
Private Const MAX_PATH = 260

#If False Then
Public Enum ShellSpecialFolderConstants
  ssfDESKTOP = &H0                   ' <Desktop>
  ssfPROGRAMS = &H2                  ' Startmenü\Programme
  ssfPERSONAL = &H5                  ' Eigene Dateien
  ssfFAVORITES = &H6                 ' <Benutzer>\Favoriten
  ssfSTARTUP = &H7                   ' Startmenü\Programme\Autostart
  ssfRECENT = &H8                    ' <Benutzer>\Recent
  ssfSENDTO = &H9                    ' <Benutzer>\SendTo
  ssfSTARTMENU = &HB                 ' <Benutzer>\Startmenü
  ssfDESKTOPDIRECTORY = &H10         ' <Benutzer>\Desktop
  ssfNETHOOD = &H13                  ' <Benutzer>\Netzwerkumgebung
  ssfFONTS = &H14                    ' Windows\Fonts
  ssfTEMPLATES = &H15                ' <Benutzer>\Vorlagen
  ssfCOMMONSTARTMENU = &H16          ' All Users\Startmenü
  ssfCOMMONPROGRAMS = &H17           ' All Users\Startmenü\Programme
  ssfCOMMONSTARTUP = &H18            ' All Users\Startmenü\Autostart
  ssfCOMMONDESKTOPDIRECTORY = &H19   ' All Users\Desktop
  ssfAPPDATA = &H1A                  ' <Benutzer>\Anwendungsdaten
  ssfPRINTHOOD = &H1B                ' <Benutzer>\Druckumgebung
  ssfCOOKIES = &H21                  ' <Benutzer>\Cookies
  ssfHISTORY = &H22                  ' <Benutzer>\Lokale Einstell.\Verlauf
  ssfCOMMONTEMPLATES = &H2D          ' All Users\Vorlagen
  ssfCOMMONDOCUMENTS = &H2E          ' All Users\Dokumente
End Enum
#End If

' für fuehraus
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd&, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd&)
Private Enum enSW
    SW_HIDE
    SW_SHOWNORMAL
    SW_SHOWMINIMIZED
    SW_MAXIMIZE
    SW_SHOWNOACTIVATE
    SW_SHOW
    SW_MINIMIZE
    SW_SHOWMINNOACTIVE
    SW_SHOWNA
    SW_RESTORE
    SW_SHOWDEFAULT
    SW_FORCEMINIMIZE
End Enum
' Datei ist keine Win32 Anwendung
Const ERROR_BAD_FORMAT = 11&
' Zugriff verweigert
Const SE_ERR_ACCESSDENIED = 5
' Datei-Assoziation ist unvollständig
Const SE_ERR_ASSOCINCOMPLETE = 27
' DDE ist nicht bereit
Const SE_ERR_DDEBUSY = 30
' DDE-Vorgang gescheitert
Const SE_ERR_DDEFAIL = 29
' DDE-Zeitlimit wurde erreicht
Const SE_ERR_DDETIMEOUT = 28
' benötigte DLL wurde nicht gefunden
Const SE_ERR_DLLNOTFOUND = 32
' Datei wurde nicht gefunden
Const SE_ERR_FNF = 2
' Datei ist nicht Assoziiert
Const SE_ERR_NOASSOC = 31
' Nicht genügend Speicher
Const SE_ERR_OOM = 8
' Pfad wurde nicht gefunden
Const SE_ERR_PNF = 3
' Sharing-Verletzung
Const SE_ERR_SHARE = 26

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Private Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd&)
Private Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Const WAIT_ABANDONED = &H80
Const WAIT_TIMEOUT = &H102
Const WAIT_OBJECT_0 = &H0
Const WAIT_FAILED = &HFFFFFFFF

Public ProgVerz$ ' c:\program files / c:\program files (x86)
Public ProgVerzO$ ' (progverz)
Public AppVerz$ ' localappdata / appdata
Public Const LiName = "linux1", LiServer$ = "\\" & LiName & "\" ' \\linux1\
Public uVerz$, pVerz$, vVerz$, plzVz$, tVerz$, xVerz$, zVerz$

'für FindProcessID
Private Const TH32CS_SNAPPROCESS        As Long = &H2&
Private Type PROCESSENTRY32 ' Prozesseintrag
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID     As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH ' = 260
End Type
Private Declare Function CreateToolhelpSnapshot& Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags&, ByVal lProcessID&)
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    
#If mitGetSpecialFolder = 1 Then
Public Function GetSpecialFolder(ByVal Folder As ShellSpecialFolderConstants) As String
  Dim tIIDL   As ITEMIDLIST
  Dim strPath As String
  
  If SHGetSpecialFolderLocation(0, Folder, tIIDL) = S_OK Then
    strPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(tIIDL.mkid.cb, strPath) <> 0 Then
      GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    End If
  End If
End Function
#End If
' Einen besonderen Ordner, in diesem Beispiel "Eigene Dateien", können Sie dann wie folgt ermitteln:
' Dim strPath AS String
' strPath = GetSpecialFolder(ssfPERSONAL)

#If False Then
' Klasse clsOperSystem.cls und Modul modTrimStr nötig, geht aber nur bis Windows 7
Function getver7()
Dim objOS As cOperSystem  ' Instantiate class object
Set objOS = New cOperSystem
End Function
#End If


'Function environ(Name$) 'dürfte wohl ersetzbar sein durch environ
'    Dim Buffer$
'    Dim l&
'    ON Error GoTo fehler
'    l = 256
'    Buffer = String$(l, vbNullChar)
'
'    l = environonmentVariable(Name, Buffer, l)
'
'    IF l <> 0 THEN
'        Buffer = LEFT(Buffer, l)
'        environ = Buffer
'    Else
'        environ = vns
'    END IF
'    Exit Function
'fehler:
'Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIGNORE, "Aufgefangener Fehler in environ/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' environ

Function SetEnvir(name, Wert)
 On Error GoTo fehler
 Call SetEnvironmentVariable(name, Wert)
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetEnvir/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetEnvir

#If maxWin7 Then
Public Function GetOSVersion() As WindowsVersion
' Konstanten
  Const VER_PLATFORM_WIN32s As Long = 0&
  Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
  Const VER_PLATFORM_WIN32_NT As Long = 2&

' Um zu testen, ob XP Home oder Professional verwendet wird.
' Weitere Informationen gibt es unter
' http://msdn.microsoft.com/library/en-us/sysinfo/base/getversionex.asp
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ ->
'     sysinfo/base/osversioninfoex_str.asp

  Const VER_SUITE_PERSONAL As Long = &H200&
 
' Private Variablen
  Static m_bAlreadyGot As Boolean
  Static m_OsVersion As WindowsVersion

' WinAPI

    Dim OsVersInfoEx As OSVERSIONINFOEX
    Dim OsVersInfo As OSVERSIONINFO
    
    On Error GoTo fehler
    
    If m_bAlreadyGot Then
        GetOSVersion = m_OsVersion
        Exit Function
    End If
    
    ' Zuerst nehmen wir nur die kleinere Struktur um sicherzugehen,
    ' dass Windows 95 und Konsorten auch damit klarkommen
    OsVersInfo.dwOSVersionInfoSize = Len(OsVersInfo)
    
    If GetVersionEx1(OsVersInfo) = 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
        MsgBox "Das Betriebssystem konnte nicht korrekt erkannt " & _
        "werden:" & _
                vbCrLf & "Fehler im API-Aufruf"
        
        m_OsVersion = WIN_OLD
        Exit Function
    End If
        
    With OsVersInfo
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32s
                m_OsVersion = WIN_OLD
            Case VER_PLATFORM_WIN32_WINDOWS
                Select Case .dwMinorVersion
                    Case 0
                        m_OsVersion = WIN_95
                    Case 10
                        m_OsVersion = WIN_98
                    Case 90
                        m_OsVersion = WIN_ME
                End Select
            Case VER_PLATFORM_WIN32_NT
                Select Case .dwMajorVersion
                    Case 3
                        m_OsVersion = WIN_NT_3x
                    Case 4
                        m_OsVersion = win_nt_4x
                    Case 5
                        Select Case .dwMinorVersion
                            Case 0
                                m_OsVersion = win_2k
                            Case 1
                                
' Es handelt sich um Windows XP. Um zu erfahren, ob das verwendete
' Produkt eine Home-Edition ist, erfragen wir die Version erneut und
' empfangen dieses Mal die komplette Liste
                                
                                OsVersInfoEx.dwOSVersionInfoSize = Len(OsVersInfoEx)
                                If GetVersionEx2(OsVersInfoEx) = 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
                                    MsgBox "Das Betriebssystem konnte nicht korrekt erkannt werden:" & _
                                        vbCrLf & "Fehler im API-Aufruf"
                                    m_OsVersion = win_xp
                                    Exit Function
                                End If
                                If (OsVersInfoEx.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL Then
                                    m_OsVersion = win_xP_home
                                Else
                                    m_OsVersion = win_xp
                                End If
                            Case 2
                                m_OsVersion = WIN_2003
                        End Select
                    Case Else
                        m_OsVersion = .dwMajorVersion + 6 ' getestet: Windows 8
                End Select
        End Select
    End With
    GetOSVersion = m_OsVersion
    m_bAlreadyGot = True
    Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetOSVersion/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetOSVersion

#Else

' nötig, da ab Windows 8 die Funktion GetVersionEx nur nach Manifest rausgibt
Function GetOSVersion&()
'Windows XP=5 (5.1.6300), Windows 7 = 6 (6.1.7601), 8.1 = 6 (6.3.9600), Windows 10=10 (10.0.10240)
Dim strComputer$
Dim objWMIService
Dim colOperatingSystems
Dim objOperatingSystem
On Error GoTo fehler
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOperatingSystem In colOperatingSystems
'    syscmd 4, objOperatingSystem.Caption & "  " & objOperatingSystem.Version
    GetOSVersion = Left$(objOperatingSystem.Version, InStr(objOperatingSystem.Version, ".") - 1)
    Exit Function
Next
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetOSVersion/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetOSVersion

#End If

Public Function HostByName(name As String, Optional x As Integer = 0) As String
    Dim MemIp() As Byte
    Dim Y As Integer
    Dim HostDeAddress As Long, HostIp As Long
    Dim IpAddress As String
    Dim Host As HostDeType
    On Error GoTo fehler
    Call InitSockets
    HostDeAddress = gethostbyname(name)
    Call CleanSockets
    If HostDeAddress = 0 Then
        HostByName = ""
        Exit Function
    End If
    
    Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
    
    For Y = 0 To x
        Call RtlMoveMemory(HostIp, Host.haddrlist + 4 * Y, 4)
        If HostIp = 0 Then
            HostByName = ""
            Exit Function
        End If
    Next Y
    
    ReDim MemIp(1 To Host.hlength)
    Call RtlMoveMemory(MemIp(1), HostIp, Host.hlength)
    
    IpAddress = ""
    
    For Y = 1 To Host.hlength
        IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
    Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HostByName/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' HostByName

Sub InitSockets()
    Dim Result As Integer
    Dim LoBy As Integer, HiBy As Integer
    Dim SocketData As WinSocketDataType
    On Error GoTo fehler
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
        Call MsgBox("'winsock.dll' antwortet nicht!")
        ProgEnde
    End If
 Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in InitSockets/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub 'InitSockets

Public Sub CleanSockets()
    Dim Result As Long
    On Error GoTo fehler
    Result = WSACleanup()
    If Result <> 0 Then
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
        Call MsgBox("Socket Error " & Trim$(CStr(Result)) & _
                " in Prozedur 'CleanSockets' aufgetreten !")
        ProgEnde
    End If
 Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CleanSockets/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' exit sub

#If ohnewsh = 0 Then
' nirgends aufgerufen
Function VerzPrüfneu$(ByVal Verz$)
 Dim Bstd$(), i%, j%, k%, tStr$, FNr&
 Dim FSO As New FileSystemObject
 On Error GoTo fehler
' IF FSO Is Nothing THEN SET FSO = CreateObject("Scripting.FileSystemObject")
 FNr = 0
 VerzPrüfneu = FSO.GetAbsolutePathName(Verz)
 FNr = 1
 Bstd = Split(VerzPrüfneu, "\")
 FNr = 2
 For j = 0 To UBound(Bstd) - 1 'da Dateiname übergeben
  tStr = ""
  FNr = 3
  For i = 0 To j
   tStr = tStr + IIf(i = 0, "", "\") + Bstd(i)
  Next i
  FNr = 4
  If LenB(tStr) <> 0 Then
   tStr = LokPfad(tStr)
   If Not FSO.FolderExists(tStr) Then
    If Not (Left$(tStr, 2) = "\\" And InStrB(Mid$(tStr, 3), "\") = 0) Then
     FNr = 5
     Call FSO.CreateFolder(tStr)
    End If
   End If
   VerzPrüfneu = tStr
  End If
 Next
Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & " tstr = " & tStr & ", j:" & j & ", ErrNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VerzPrüfneu/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' VerzPrüf(ByVal Verz$)

' aufgerufen in do_Datenbank_Aufruf_Click, Formular.do_Form_Open, VerzPrüfneu
Public Function LokPfad$(Pfad)
Dim CName$(1), IP$, i%, j%, teststr$, neuPfad$
On Error GoTo fehler
LokPfad = Pfad
CName(0) = CptName
CName(1) = HostByName(CptName)
ListFreigaben
For i = 1 To FZ - 1
 For j = 0 To 1
  teststr = LCase$("\\" + CName(j) + "\" + FNam(i))
  If InStr(1, Pfad, Trim$(teststr), vbTextCompare) = 1 Then
   LokPfad = FInh(i) + Mid$(Pfad, Len(Trim$(teststr)) + 1)
   Exit For
  End If
 Next j
Next i
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LokPfad/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
 Exit Function
End Function ' LokPfad

' aufgerufen in LokPfad
Function ListFreigaben()
' Freigebene Verzeichnisse des aktuellen PCs auflisten
Dim nr&, erg$, Inhalt$
On Error GoTo fehler
nr = 0
FZ = 0
Do
erg = ReadRegistryGetVALUES(&H80000002, "SYSTEM\ControlSet001\Services\lanmanserver\Shares", nr, Inhalt)
'erg = ReadRegistryGetVALUES(&H80000002, "SYSTEM\CurrentControlSet\Services\lanmanserver\Shares", Nr, Inhalt)
 If LenB(Inhalt) <> 0 Then
  If Mid$(Inhalt, 2, 1) = ":" Then
   ReDim Preserve FNam(FZ)
   ReDim Preserve FInh(FZ)
   FNam(FZ) = erg
   FInh(FZ) = Inhalt
   FZ = FZ + 1
'   Debug.Print erg
'   Debug.Print Inhalt
  End If
 End If
 nr = nr + 1
Loop Until LenB(Trim$(erg)) = 0
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ListFreiGaben/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ListFreigaben
#End If

Function VerzPrüf(ByVal Verz$, Optional obaggr%)
 Dim Bstd$(), i%, j%, k%, tStr$
 Dim FSO As New FileSystemObject
 On Error GoTo fehler
' IF FSO Is Nothing THEN SET FSO = CreateObject("Scripting.FileSystemObject")
 VerzPrüf = FSO.GetAbsolutePathName(Verz)
 Bstd = Split(VerzPrüf, "\")
 For j = 0 To UBound(Bstd)
'  tStr = vNS
'  For i = 0 To j
'   tStr = tStr + IIf(i = 0, vNS, "\") + Bstd(i)
'  Next i
  tStr = tStr + IIf(j = 0, "", "\") + Bstd(j)
  If j > 0 Then machOrdner tStr, obaggr
 Next
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VerzPrüf/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' VerzPrüf(ByVal Verz$)

Function VerzPrüfAlt(Verz$)
 Dim Spli$(), i%, Zus$
 Dim FSO As New FileSystemObject
 On Error GoTo fehler
' IF FSO Is Nothing THEN SET FSO = CreateObject("Scripting.FileSystemObject")
 Spli = Split(Verz, "\")
 If LenB(Verz) = 0 Or Verz = "." Or Verz = ".." Then Exit Function
 If UBound(Spli) > -1 And LenB(Spli(0)) = 0 Then i = 1
 If UBound(Spli) > 0 And LenB(Spli(1)) = 0 Then Zus = "\"
 For i = i To UBound(Spli)
  If LenB(Spli(i)) <> 0 Then
   Zus = Zus + IIf(LenB(Zus) = 0, "", "\") + Spli(i)
   If Not FSO.FolderExists(Zus) Then
    If Not (Left$(Zus, 2) = "\\" And InStr(3, Zus, "\") = 0) Then
     Call FSO.CreateFolder(Zus)
    End If
   End If
  End If
 Next i
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VerzPrüfAlt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' VerzPrüfAlt

' in meld, Konstanten, doHAAAkt
Public Sub SetProgV()
 If LenB(ProgVerz) = 0 Then
  If WV = 0 Then WV = GetOSVersion
  If WV < win_vista Then
   ProgVerz = Environ("programfiles")
   AppVerz = Environ("appdata")
  Else ' WV < win_vista Then
   ProgVerzO = Environ("programw6432")
   If ProgVerzO = "" Then ProgVerzO = Environ("programfiles")
   ProgVerz = Environ("programfiles(x86)")
   If LenB(ProgVerz) = 0 Then ProgVerz = ProgVerzO ' 32-bit-Systeme
   AppVerz = Environ("localappdata")
  End If ' WV < win_vista Then else
  uVerz = IIf(FSO.FolderExists("u:"), "u:", LiServer & "Daten\eigene Dateien") & "\"
  pVerz = IIf(FSO.FolderExists("p:"), "p:", LiServer & "Daten\Patientendokumente") & "\"
  vVerz = IIf(FSO.FolderExists("v:"), "v:", LiServer & "Daten\down") & "\"
  tVerz = IIf(FSO.FolderExists("t:"), "t:", LiServer & "Daten\shome\gerald") & "\"
  xVerz = IIf(FSO.FolderExists("x:"), "x:", LiServer & "turbomed") & "\"
  zVerz = IIf(FSO.FolderExists("z:"), "z:", LiServer & "Daten") & "\"
  plzVz = pVerz & "plz\"
  ProgVerz = Environ("programfiles") ' ab 8.1.24, zuvor "c:\programme"
  If Right$(ProgVerz, 1) <> "\" Then ProgVerz = ProgVerz & "\"
 End If
End Sub ' SetProgV()

' in doVorhandene, tubriefStandalone, GetVorDat, Epikrise
Function meld(Text$, Optional obStumm%)
 Dim MeldDatei$
 On Error GoTo fehler
 Call SetProgV
 MeldDatei = uVerz & "meldung " & Format$(Now, "dd.mm.yy hh.mm.ss")
 Open MeldDatei For Output As #299
 Print #299, Text
 Close #299
 If Not obStumm Then zeigan MeldDatei
  ' Call SuSh(ProgVerz & "\notepad++\notepad++ """ & MeldDatei & """", 0, , 0, 1)
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Meld/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' meld

Public Function doUmwfSQL$(q$, obmy As Boolean, Optional mittrim% = True)
 If mittrim Then doUmwfSQL = Trim$(q) Else doUmwfSQL = q
 If InStrB(doUmwfSQL, "\") <> 0 Then
  doUmwfSQL = REPLACE$(doUmwfSQL, "\", "\\")
 End If
 If InStrB(doUmwfSQL, "'") <> 0 Then
  doUmwfSQL = REPLACE$(doUmwfSQL, "'", IIf(obmy, "\'", "''"))
 End If
 If InStrB(doUmwfSQL, Chr$(0)) <> 0 Then ' aus umw in tabübertr
  doUmwfSQL = REPLACE$(doUmwfSQL, Chr$(0), "")
 End If
' IF InStrB(doUmwfSQL, "¿") <> 0 THEN
'  doUmwfSQL = replace$(doUmwfSQL, "¿", "\¿")
' END IF
 If InStrB(doUmwfSQL, """") <> 0 Then ' aus doCopyDaten in MachDatenbank
  If obmy Then
   doUmwfSQL = REPLACE$(doUmwfSQL, """", "\""")
  End If
 End If
End Function ' doUmwfSQL

Function DateiVergleichen%(D1$, D2$)
 Dim p1$, p2$
 On Error Resume Next
 Open D1 For Input As #355 Len = 1000
 If Err.Number <> 0 Then Exit Function
 Open D2 For Input As #356 Len = 1000
 If Err.Number <> 0 Then Exit Function
 On Error GoTo fehler
 DateiVergleichen = True
 Do While Not EOF(355) And Not EOF(356)
'  P1 = input(100, #355)
  Input #355, p1
'  P2 = input(100, #356)
  Input #356, p2
  If p1 <> p2 Then
   DateiVergleichen = False
   Exit Do
  End If
 Loop
 If EOF(355) <> EOF(356) Then DateiVergleichen = False
 Close #355
 Close #356
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DateiVergleichen/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DateiVergleichen


'Public FUNCTION testAktiv%()
' Dim col As New Collection, AppZahl%, i%
' ON Error GoTo fehler
' SET col = GetProcessCollection
' For i = 1 To col.Count
'  Debug.Print col(i), App.EXEName
'  IF InStrB(col(i), "|" & ProcIDFromWnd(App.ThreadID)) = 0 THEN ' nicht aktuelle Anwendung
'   IF InStrB(col(i), App.EXEName) <> 0 THEN
'    AppZahl = AppZahl + 1
'    IF AppZahl > 0 THEN
''   MsgBox "Programm wird bereits ausgeführt!"
'     testAktiv = True
'     Exit For
'    END IF
'   ElseIf InStrB(UCase$(col(i)), "NVINI") <> 0 OR InStrB(UCase$(col(i)), "NVERB") <> 0 THEN
'    testAktiv = True
'    Exit For
'   END IF
'  END IF
' Next i
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in testAktiv/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' testAktiv

Public Sub Pause(Millisekunden As Long)
 Sleep Millisekunden
End Sub ' Pause

Function TütSoundkarte()
' Const MB_ICONASTERISK = &H40& ' Warnung
' Const MB_ICONEXCLAMATION = &H30& ' Hinweis
' Const MB_ICONHAND = &H10& ' Info
' Const MB_ICONQUESTION = &H20& ' Frage
' Const MB_OK = &H0& ' Standard Sound
' Dim Retval AS Long
' Retval = MessageBeep(MB_ICONQUESTION)
' IF Retval = 0 THEN
'  Debug.Print "MessageBeep ist gescheitert"
' END IF
Const SND_ALIAS = &H10000 ' der angegebene Name muss ein Eintrag aus der
' WIn.ini unter [Sounds] sein
Const SND_ALIAS_ID = &H110000 ' der angegebene Name muss ein Key aus der
' Win.ini unter [Sounds] sein
Const SND_APPLICATION = &H80 ' der angegebene Name ist ein Ereignissound
Const SND_ASYNC = &H1 ' Stoppt die Wiedergabe aller Sounddateien, um diese abzuspielen
Const SND_FILENAME = &H20000 ' der angegebene Name ist ein Pfad zu einer
' Wave-Datei
Const SND_LOOP = &H8 ' wiederholt unendlich oft die Wiedergabe
Const SND_MEMORY = &H4 ' der angegebene Name ist ein Byte Array mit den
' Wave-Datei-Daten
Const SND_NODEFAULT = &H2 ' spielt keinen Standardsound ab wenn die
' angegebene Datei nicht gefunden wird
Const SND_NOSTOP = &H10 '  stoppt keine momentan laufenden Sounds
Const SND_NOWAIT = &H2000 ' wartet nicht auf das Beenden des laufenden Sounds
' um dann den angegebenen Sound abzuspielen
Const SND_PURGE = &H40 ' stoppt die unendliche Wiedergabe der Sounds, die mit
' SND_LOOP eingeleitet wurde
Const SND_RESOURCE = &H40004 ' der angegebene Name ist der Name einer
' Ressource in der sich die Wave-Datei befindet, hierfür muss hModule das
' Modul-Handle der Anwendung bekommen, die die Ressource besitzt
Const SND_SYNC = &H0 ' die Funktion kehrt erst nach Beenden der Wiedergabe
' des Sounds zurück

  Dim RetVal As Long, FFile As Long, MemWav() As Byte
 
  ' öffnen der Datei und übertragen in MemWav
  FFile = FreeFile
  Open Environ("windir") & "\media\" & "Windows Navigation Start.wav" For Binary As FFile
  ReDim MemWav(LOF(FFile))
  Get FFile, , MemWav()
  Close FFile
 
  ' abspielen von Tata aus dem Speicher und auf Beenden des Sounds warten
  RetVal = PlaySound(MemWav(0), 0&, SND_MEMORY Or SND_NODEFAULT)

End Function ' TütSoundkarte

Public Function lesetest()
 Dim zln$(), zle$, zlnr&, maxnr&
 Open uVerz & "tmexport\20221230.BDT" For Input As #23
 Do While Not EOF(23)
  If zlnr = maxnr Then
   maxnr = maxnr + 1000
   ReDim Preserve zln(maxnr)
  End If
  Line Input #23, zln(zlnr)
  zlnr = zlnr + 1
 Loop
 Close #23
End Function ' lesetest

#If zutesten Then
Public Function lesetest2()
 Dim zln$(), zle$, zlnr&, maxnr&, pos&
 Open uVerz & "tmexport\20221230.BDT" For Input As #23
 
 Do While Not EOF(23)
  Line Input #23, zle
  pos = Seek(23)
  Debug.Print pos, zle
  zlnr = zlnr + 1
  Seek #23, 14
 Loop
 Close #23
 Debug.Print zlnr
End Function ' lesetest
#End If

#If zutesten Then
' zu groß für:
Public Function Lesetest3(ByVal sFilename As String)
  Dim F As Integer
  Dim sInhalt As String
  ' Prüfen, ob Datei existiert
  If Dir$(sFilename, vbNormal) <> "" Then
    ' Datei im Binärmodus öffnen
    F = FreeFile: Open sFilename For Binary As #F
    ' Größe ermitteln und Variable entsprechend
    ' mit Leerzeichen füllen
    sInhalt = Space$(FSO.GetFile(sFilename).size)
    ' Gesamten Inhalt in einem "Rutsch" einlesen
    Get #F, , sInhalt
    ' Datei schliessen
    Close #F
  End If
End Function ' lesetest3
#End If

Public Sub machOrdner(tStr$, Optional obaggr%)
 Dim gibts%, runde%
 On Error GoTo fehler
  If WV = 0 Then WV = GetOSVersion
  If LenB(tStr) <> 0 Then
    If Right$(tStr, 1) = "\" Then tStr = Left$(tStr, Len(tStr) - 1)
    For runde = 1 To 2
     If DirExists(tStr) Then
      gibts = True
     ElseIf FSO.FolderExists(tStr) Or Not InStrB(Mid$(tStr, 3), "\") <> 0 Then
      gibts = True
     End If
     If Not gibts Or obaggr Then
      If WV < win_vista Then
       Call FSO.CreateFolder(tStr)
      Else
       rufauf "cmd", "/c mkdir """ & tStr & """", IIf(runde = 1, 0, 2), , , 0
      End If
     Else
      Exit For
     End If
    Next runde
'      Shell (vVerz & doalsad & acceu & ap1 & ap2 & " cmd /e:on /c mkdir " & Chr$(34) & tStr & Chr$(34))
'      SuSh "cmd /e:on /c mkdir " & Chr$(34) & tStr & Chr$(34), 2, , 0
  End If
  Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in machOrdner/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' machOrdner

#If True Then
Public Function KopDat%(q$, z$, Optional WV As WindowsVersion)
 Static iWV As WindowsVersion, ErrNr&
 Static pruefe%
 Dim Datei$, obcopy&
 Dim ZDat$
 On Error GoTo fehler
 Datei = Right$(q, Len(q) - InStrRev(q, "\"))
 If Right$(z, 1) = "\" Then ZDat = z & Datei Else ZDat = z
 On Error Resume Next
 If FileExists(q) Then 'Or fileexists(q) THEN
  FileCopy q, z
  If Not FileExists(ZDat) Then obcopy = True Else If FSO.GetFile(ZDat).size <> FSO.GetFile(q).size Or FileDateTime(ZDat) <> FileDateTime(q) Then obcopy = True Else obcopy = False ' 21.8.21 statt FileLen(
  If obcopy Then
   FSO.CopyFile q, z, True
   If Not FileExists(ZDat) Then obcopy = True Else If FSO.GetFile(ZDat).size <> FSO.GetFile(q).size Or FileDateTime(ZDat) <> FileDateTime(q) Then obcopy = True Else obcopy = False
   If obcopy Then
    If iWV = 0 Then If WV = 0 Then iWV = GetOSVersion Else iWV = WV
    If iWV >= win_vista Then
    On Error GoTo fehler
     If Not pruefe Then
      Call adminaktiv
      pruefe = True
     End If
     Dim UPZDat$
     UPZDat = Environ("userprofile") & "\" & Datei
     If FileExists(UPZDat) Then
       On Error Resume Next
       Kill UPZDat
       On Error GoTo fehler
     End If
     If Not FileExists(UPZDat) Then
'      On Error Resume Next
      FileCopy q, UPZDat
      On Error GoTo fehler
      If FileExists(UPZDat) Then
       rufauf "cmd", "/c move """ & UPZDat & """ """ & z & """", 0, , , 0
       If Not FileExists(ZDat) Then obcopy = True Else If FSO.GetFile(ZDat).size <> FSO.GetFile(q).size Or FileDateTime(ZDat) <> FileDateTime(q) Then obcopy = True Else obcopy = False
       If obcopy Then
        rufauf "cmd", "/c move """ & UPZDat & """ """ & z & """", 2, , , 3
        If Not FileExists(ZDat) Then obcopy = True Else If FSO.GetFile(ZDat).size <> FSO.GetFile(q).size Or FileDateTime(ZDat) <> FileDateTime(q) Then obcopy = True Else obcopy = False
        If obcopy Then
         rufauf "cmd", "/c move """ & UPZDat & """ """ & z & """", 1, , , 3
        End If
       End If
      End If
     End If
    End If
   End If
   If Not FileExists(ZDat) Then obcopy = True Else If FSO.GetFile(ZDat).size <> FSO.GetFile(q).size Or FileDateTime(ZDat) <> FileDateTime(q) Then obcopy = True Else obcopy = False
   If obcopy Then
    Debug.Print "Fehler beim Erstellen von: " & ZDat
'    syscmd 4, "Fehler beim Erstellen von: " & ZDat
   Else
    KopDat = True
   End If
  End If
 End If
 Exit Function
fehler:
 Select Case MsgBox("Fpos: " & FPos & " ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KopDat/" & App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' KopDat

#Else
Public Sub KopDat(q$, z$, Optional WV As WindowsVersion)
 Static userprof$
 If userprof = "" Then userprof = Environ("userprofile")
 Dim dname0$
 Dim fdt As Date
 Dim size&
 On Error GoTo fehler
 Static iWV As WindowsVersion
 Static pruefe%
 Dim Datei$
 Datei = Right$(q, Len(q) - InStrRev(q, "\"))
 If iWV = 0 Then If WV = 0 Then iWV = GetOSVersion Else iWV = WV
 If iWV < win_vista Then
  If FileExists(z & Datei) Then
   fdt = FileDateTime(z & Datei)
   size = FSO.GetFile(z & Datei).size
  End If
  On Error Resume Next
  FileCopy q, z
  On Error GoTo fehler
  Dim fsokop%
  If Not FileExists(z & Datei) Then
   fsokop = True
  Else
   fsokop = True
   If dname0 <> "" Then
    If FSO.GetFile(q).size = size Then
     If FileDateTime(q) = fdt Then
      fsokop = False
     End If
    End If
   End If
  End If
  If fsokop Then
'   ON Error Resume Next
   FSO.CopyFile q, z
'   Dim copyErr%
'   copyErr = Err.Number
'   ON Error GoTo fehler
'   IF copyErr THEN
'    FSO.CopyFile REPLACE(LCase$(q), zVerz, "z:\"), z
'   END IF
  End If
 Else
  If Not pruefe Then
    Call adminaktiv
    pruefe = True
  End If
  #If False Then
   ShellaW vVerz & doalsAd & acceu & ap1 & ap2 & " cmd /e:on /c xcopy /y /h /r " & Chr$(34) & q & Chr$(34) & " " & Chr$(34) & z & Chr$(34), vbHide, , 10000
'   WarteAufNicht "Administrator: C:\Windows\System32\cmd.exe", 15
  #Else
'  ShellaW vverz & doalsad & acceu & ap1 & ap2 & " cmd /e:on /c xcopy /d /h /r /y " & Chr$(34) & q & Chr$(34) & " " & Chr$(34) & userprof & "\" & Chr$(34), vbHide, , 10000
'  ShellaW "cmd /c xcopy /d /h /r /y " & Chr$(34) & q & Chr$(34) & " " & Chr$(34) & userprof & "\" & Chr$(34), vbHide, , 10000
'  SuSh "cmd /c xcopy /d /h /r /y " & Chr$(34) & q & Chr$(34) & " " & Chr$(34) & userprof & "\" & Chr$(34), 3, , 0
   rufauf "cmd", "/c xcopy /d /h /r /y """ & q & """ """ & userprof & "\""", , , , 0
'  FileCopy q, UP & "\" & Datei
'  Shell (vverz & doalsad & acceu & ap1 & ap2 & " cmd /e:on /c xcopy /d /h /r /y " & Chr$(34) & userprof & "\" & Datei & Chr$(34) & " " & Chr$(34) & z & Chr$(34))
  DoEvents
'  SuSh "cmd /c xcopy /d /h /r /y " & Chr$(34) & userprof & "\" & Datei & Chr$(34) & " " & Chr$(34) & z & Chr$(34), 3, , 0
   rufauf "cmd", "/c xcopy /d /h /r /y """ & userprof & "\" & Datei & """ """ & z & """", , , , 0
  #End If
 End If
 If Not FileExists(z) Then ' & IIf(Right$(z, 1) = "\", Datei, "")
  If Not FSO.FolderExists(z) Then  ' 21.8.15, noetig fuer 'c:\program files'
'   ShellaW vverz & doalsad & acceu & ap1 & ap2 & " cmd /e:on /c xcopy /d /h /r " & Chr$(34) & userprof & "\" & Datei & Chr$(34) & " " & Chr$(34) & Left$(z, InStrRev(z, "\")) & Chr$(34), vbHide, , 10000
'   SuSh "xcopy /d /h /r " & Chr$(34) & userprof & "\" & Datei & Chr$(34) & " " & Chr$(34) & Left$(z, InStrRev(z, "\")) & Chr$(34), 3, , 0
    rufauf "xcopy", "/d /h /r """ & userprof & "\" & Datei & """ """ & Left$(z, InStrRev(z, "\")) & """", 0, , 0, 0
'   Shell vverz & doalsad & acceu & ap1 & ap2 & " cmd /e:on /c del " & Chr$(34) & userprof & "\" & Datei & Chr$(34), vbHide
'   SuSh "cmd /e:on /c del " & Chr$(34) & userprof & "\" & Datei & Chr$(34), 1, , 0
   rufauf "cmd", "/e:on /c del """ & userprof & "\" & Datei & """", , , 0, 0
   If Not FileExists(z) Then ' & IIf(Right$(z, 1) = "\", Datei, "")
    Debug.Print "Fehler beim Erstellen von: " & z
    syscmd 4, "Fehler beim Erstellen von: " & z
   End If ' Not FileExists(z) Then ' & IIf(Right$(z, 1) = "\", Datei, "")
  End If ' Not FSO.FolderExists(z) Then  ' 21.8.15, noetig fuer 'c:\program files'
 End If ' Not FileExists(z) Then ' & IIf(Right$(z, 1) = "\", Datei, "")
 DoEvents
 Exit Sub
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("Fehler beim Kopieren von '" & q & "'" & vbCrLf & " nach '" & z & "':" & vbCrLf & "FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in KopDat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' KopDat
#End If

Public Function FindProcessID(ByVal pExename As String) As Long
    Dim ProcessID As Long, hSnapShot As Long
    Dim uProcess As PROCESSENTRY32, rProcessFound As Long
    Dim pos As Integer, szExename As String
    ' Create snapshot of current processes
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ' Check IF snapshot is valid
    If hSnapShot = -1 Then
        Exit Function
    End If
    'Initialize uProcess with correct size
    uProcess.dwSize = Len(uProcess)
    'Start looping through processes
    rProcessFound = ProcessFirst(hSnapShot, uProcess)
    Do While rProcessFound
        pos = InStr(1, uProcess.szExeFile, vbNullChar)
        If pos Then
            szExename = Left$(uProcess.szExeFile, pos - 1)
        End If
        If LCase$(szExename) = LCase$(pExename) Then
            'Found it
            ProcessID = uProcess.th32ProcessID
            Exit Do
          Else
            'Wrong, so continue looping
            rProcessFound = ProcessNext(hSnapShot, uProcess)
        End If
    Loop
    CloseHandle hSnapShot
    FindProcessID = ProcessID
End Function ' FindProcessID


' aus fuehraus in fürIcon
' millis: -1: warte ewig, 0: warte gar nicht, >0: Wartezeit in Millisekunden
' alsAdm: 0: nein, 1, -1: ja, mit Prompt, 2: mit psexec
Public Function rufauf&(Datei$, Optional Para$, Optional alsAdm%, Optional vz$, Optional dwmillis& = 60000, Optional fengroe% = 1, Optional obkill&)
'Public FUNCTION SuSh&(ByVal App$, Optional alsAdm&, Optional ByVal WorkDir$, Optional dwmillis& = 10000, _
    Optional ByVal start_size& = SW_HIDE, Optional ByVal Priority_Class& = NORMAL_PRIORITY_CLASS, _
    Optional obkill&)
  Dim hierdir$
  Dim iru%
  Dim RetVal&, lHwnd&, pid&
  Dim hDatei$, hPara$, FMld$
  syscmd 4, "Rufe auf: " & Datei & " " & Para & " in " & vz
  rufauf = False
  If hierdir = vbNullString Then hierdir = Environ("userprofile")
  If vz <> vbNullString Then hierdir = vz
'  For iru = 2 To 1 Step -1
   If alsAdm = 2 Then
    hPara = acceu & ap1 & ap2 & " " & Datei & " " & Para
    hDatei = vVerz & doalsAd
   Else
    hPara = Para
    hDatei = Datei
   End If
   RetVal = ShellExecute(0, IIf(Abs(alsAdm) = 1, "runas", vbNullString), hDatei, hPara, hierdir, fengroe)
'   IF RetVal <> 0 AND alsAdm = 2 THEN alsAdm = 1 ELSE Exit For
'  Next iru
  If RetVal >= 0 And RetVal < 33 Then FMld = "Shellexecute " & hDatei & " " & hPara & ", vz: " & vz & ", alsAdm: " & alsAdm & ": "
  Select Case RetVal
    Case SE_ERR_NOASSOC
      MsgBox FMld & "Datei ist nicht Assizoiert", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_PNF
      MsgBox FMld & "Pfad wurde nicht gefunden", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_FNF
      MsgBox FMld & "Datei wurde nicht gefunden", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_OOM
      MsgBox FMld & "Nicht genügend Speicher", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_SHARE
      MsgBox FMld & "Sharing-Verletzung", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_DLLNOTFOUND
      MsgBox FMld & "benöt. DLL nicht gefunden", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_DDETIMEOUT
      MsgBox FMld & "DDE-Zeitlimit wurde erreicht", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_DDEFAIL
      MsgBox FMld & "DDE-Vorgang gescheitert", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_DDEBUSY
      MsgBox FMld & "DDE nicht bereit", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_ASSOCINCOMPLETE
      MsgBox FMld & "Datei-Assoziation unvollständig", vbInformation, "Fehler"
      Exit Function
    Case SE_ERR_ACCESSDENIED
      MsgBox FMld & "Zugriff verweigert", vbInformation, "Fehler"
      Exit Function
    Case ERROR_BAD_FORMAT
      MsgBox FMld & "Datei keine zulässige Win32-Anwendung", vbInformation, "Fehler"
      Exit Function
    Case Is > 32 ' Handle
    Case Else
      rufauf = True
      If dwmillis <> 0 Then ' 0 = asynchron
       If alsAdm = 2 Then hDatei = FSO.GetFileName(vVerz & doalsAd) Else If InStrB(hDatei, "\") <> 0 Then hDatei = FSO.GetFileName(hDatei)
       pid = FindProcessID(hDatei)
  '    lHwnd = GetWinHandle(PID) ' funzt (zumindest hier) nicht
       If pid <> 0 Then
        Const SYNCHRONIZE = &H100000
'        Const INFINITE = &HFFFF ' -1
        lHwnd = OpenProcess(SYNCHRONIZE, 0, pid)
        If lHwnd <> 0 Then
         RetVal = WaitForSingleObject(lHwnd, dwmillis)
         If RetVal = WAIT_ABANDONED Or RetVal = WAIT_TIMEOUT Then ' nonsignaled
          If obkill <> 0 Then
           KillProcessByPID (pid)
          End If
         End If
        End If
        CloseHandle (lHwnd)
       End If
      End If
  End Select
  syscmd 5
  Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in rufauf/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' rufauf

Sub adminaktiv()
 Dim Text$, f1
 On Error GoTo 0
' Shell ("cmd.exe /c net user administrator > " & Chr$(34) & Environ("appdata") & "\obadmin.txt" & Chr$(34))
' SuSh "cmd.exe /c net user administrator > " & Chr$(34) & Environ("appdata") & "\obadmin.txt" & Chr$(34)
nochmal:
 rufauf "cmd.exe", "/c net user administrator > """ & Environ("appdata") & "\obadmin.txt" & """", , , 2000, 0
 Do While Not FileExists(Environ("appdata") & "\obadmin.txt")
 Loop
 Do While FSO.GetFile(Environ("appdata") & "\obadmin.txt").size < 1000 ' 1044 ist die Zielgröße ' 21.8.21 statt FileLen
 Loop
 Open Environ("appdata") & "\obadmin.txt" For Input As #193
 Do While Not EOF(193)
  Line Input #193, Text
  If Left$(Text, 11) = "Konto aktiv" Then
   If Text Like "*Nein*" Then
    rufauf "cmd", "/c net user Administrator " & p1 & ap2 & " /active:yes", 1, , , 0
    Close #193
    GoTo nochmal:
    MsgBox "Bitte aktivieren Sie den Administrator von cmd als Administrator mit 'net user Administrator * /active:yes' und starten Sie das Programm dann nochmal!"
    Unload FürIcon
    End
' geht nicht, da der Administrator nicht vom Nicht-Administrator aktiviert werden kann
'    Shell ("cmd.exe /c net user Administrator /active:yes")
   End If
   Exit Do
  End If
 Loop
 Close #193
 On Error Resume Next
 Kill Environ("appdata") & "\obadmin.txt"
 DoEvents
 If FileExists(Environ("appdata") & "\obadmin.txt") Then
  rufauf "cmd.exe", "/c del """ & Environ("appdata") & "\obadmin.txt""""", , , 500, 0
 End If
 If FileExists(Environ("appdata") & "\obadmin.txt") Then
  rufauf "cmd.exe", "/c del """ & Environ("appdata") & "\obadmin.txt""""", 2, , 500, 0
 End If
End Sub ' adminaktiv


