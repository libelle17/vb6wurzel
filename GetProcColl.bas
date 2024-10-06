Attribute VB_Name = "GetProcColl"
Option Explicit

' für Supershell
Const INFINITE = &HFFFF
Const STARTF_USESHOWWINDOW = &H1

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
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'Private Type SECURITY_ATTRIBUTES
'    nLength AS Long
'    lpSecurityDescriptor AS Long
'    bInheritHandle AS Long
'End Type
Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
'Private Declare FUNCTION CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName AS String, ByVal lpCommandLine AS String, lpProcessAttributes AS SECURITY_ATTRIBUTES, lpThreadAttributes AS SECURITY_ATTRIBUTES, ByVal bInheritHandles AS Long, ByVal dwCreationFlags AS Long, lpEnvironment AS Any, ByVal lpCurrentDriectory AS String, lpStartupInfo AS STARTUPINFO, lpProcessInformation AS PROCESS_INFORMATION) AS Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
' Ende Supershell

Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Dim ErrNumber&, ErrDescr$, FNr&

' -----------------------------------------------------------------
' Liste aktiver Anwendungsprozesse ermitteln
' Copyright © Mathias Schiffer 1999-2005
' -----------------------------------------------------------------
'
' KURZE FUNKTIONSBESCHREIBUNG:
'
' - Public FUNCTION GetProcessCollection() AS Collection
'   Gibt eine String-Collection zurück, deren Einträge den
'   Aufbau "Prozessname|Prozess-ID" haben.
'
' - Public FUNCTION ProcessName(ByVal CollectionString AS String) AS String
'   Extrahiert aus einem String der Collection den Prozessnamen.
'
' - Public FUNCTION ProcessHandle(ByVal CollectionString AS String) AS Long
'   Extrahiert aus einem String der Collection die Prozess-ID.
'
' - Public FUNCTION KillProcessByPID(ByVal PID AS Long) AS Boolean
'   Terminiert einen Prozess auf Basis seiner Prozess-ID. Ein Prozess
'   sollte nur in "Notfällen" terminiert werden. Datenverluste der
'   terminierten Anwendung sind nicht ausgeschlossen.
'
' -----------------------------------------------------------------
  
  
' ------------------------- DEKLARATIONEN -------------------------
  
' Deklaration notwendiger API-Funktionen:
  
' GetVersionEx dient der Erkennung des Betriebssystems:
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO)
' Toolhelp-Funktionen zur Prozessauflistung (Win9x):
Private Declare Function CreateToolhelp32Snapshot& Lib "kernel32" (ByVal dwFlags&, ByVal th32ProcessID&)
Private Declare Function Process32First& Lib "kernel32" (ByVal hSnapShot&, ByRef lppe As PROCESSENTRY32)
Private Declare Function Process32Next& Lib "kernel32" (ByVal hSnapShot&, ByRef lppe As PROCESSENTRY32)
' PSAPI-Funktionen zur Prozessauflistung (Windows NT)
Private Declare Function EnumProcesses& Lib "psapi.dll" (ByRef lpidProcess&, ByVal cb&, ByRef cbNeeded&)
Private Declare Function GetModuleFileNameEx _
  Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
  ByVal hProcess As Long, _
  ByVal hModule As Long, _
  ByVal ModuleName As String, _
  ByVal nSize As Long _
  ) As Long
Private Declare Function EnumProcessModules _
  Lib "psapi.dll" ( _
  ByVal hProcess As Long, _
  ByRef lphModule As Long, _
  ByVal cb As Long, _
  ByRef cbNeeded As Long _
  ) As Long
' Win32-API-Funktionen für Prozessmanagement
Private Declare Function OpenProcess _
  Lib "kernel32.dll" ( _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwProcId As Long _
  ) As Long
Declare Function TerminateProcess _
  Lib "kernel32" ( _
  ByVal hProcess As Long, _
  ByVal uExitCode As Long _
  ) As Long
Private Declare Function CloseHandle& Lib "kernel32.dll" (ByVal Handle&)
  
' Deklaration notwendiger Konstanter:
  
Private Const MAX_PATH                  As Long = 260
Private Const PROCESS_QUERY_INFORMATION As Long = 1024 ' &H400
Private Const PROCESS_VM_READ           As Long = 16
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_ALL_ACCESS        As Long = STANDARD_RIGHTS_REQUIRED _
                                               Or SYNCHRONIZE Or &HFFF
Private Const PROCESS_TERMINATE         As Long = &H1
Private Const TH32CS_SNAPPROCESS        As Long = &H2&
  
' Konstante für die Erkennung des Betriebssystems:
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
  
' Notwendige Typdeklarationen
  
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
  
Private Type OSVERSIONINFO ' Betriebssystemerkennung
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5
Private Const GWL_STYLE = (-16)
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const WS_MINIMIZE = &H20000000 ' Style bit 'is minimized'
Private Const obAuchTitellose% = -1
Private Const obAuchUnsichtbare% = -1
'Private Const SW_RESTORE = 9 ' Restore window
Private Const HWND_TOP = 0 ' Move to top of z-order
Private Const SWP_NOMOVE = &H2 ' Do not reposition window
Private Const SWP_NOSIZE = &H1 ' Do not re-size window
Private Const SWP_SHOWWINDOW = &H40 ' Make window visible/active
Public Const WM_CLOSE = &H10

Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, ByVal wIndx&)
Private Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd&)
Private Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd&)
Private Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hwnd&, lpdwProcessId&)
Private Declare Function GetDesktopWindow& Lib "user32" ()
Private Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Private Declare Function GetForegroundWindow& Lib "user32" ()
Private Declare Function ShowWindow& Lib "user32.dll" (ByVal hwnd&, ByVal nCmdShow&)
Private Declare Function SetWindowPos& Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub SetForegroundWindow Lib "user32" (ByVal hwnd&)
Public Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function CreateJobObject Lib "kernel32.dll" Alias "CreateJobObjectA" (lpJobAttributes As SECURITY_ATTRIBUTES, lpName As String) As Long
Public Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Public Declare Function TerminateJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Long
' Declare FUNCTION CloseHandle Lib "kernel32" (ByVal hObject AS Long) AS Long

Declare Function GetExitCodeProcess& Lib "kernel32" (ByVal hProcess&, lpExitCode&)
' Declare FUNCTION TerminateProcess Lib "kernel32" (ByVal hProcess AS Long, ByVal uExitCode AS Long) AS Long

Public TitelList As Collection, ProcIDList As Collection, hWndList As Collection
' für ShellandWaitfortermination
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
'Private Declare FUNCTION GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess AS Long, lpExitCode AS Long) AS Long
Private Declare Function timeGetTime& Lib "winmm.dll" ()
'Private Declare FUNCTION OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess AS Long, ByVal bInheritHandle AS Long, ByVal dwProcessId AS Long) AS Long
Private Const STILL_ACTIVE = &H103
'Private Const PROCESS_QUERY_INFORMATION = &H400
'Private Declare FUNCTION CloseHandle Lib "kernel32" ( _
    ByVal hObject AS Long) AS Long

Dim FPos&

Private Sub GetWindowInfo(ByVal hwnd&)
  Dim Parent&, Task&, Result&, x&, style&, Title$
  On Error GoTo fehler
  
    'Darstellung des Fensters
    style = GetWindowLong(hwnd, GWL_STYLE)
    style = style And (WS_VISIBLE Or WS_BORDER)
            
    'Titel des Fenster auslesen
    Result = GetWindowTextLength(hwnd) + 1
    Title = Space$(Result)
    Result = GetWindowText(hwnd, Title, Result)
    Title = Left$(Title, Len(Title) - 1)
    
    'In Abhängigkeit der Optionen die Ausgabe erstellen
    If (LenB(Title) <> 0 Or obAuchTitellose) And _
       (style = (WS_VISIBLE Or WS_BORDER) Or obAuchUnsichtbare) Then
       
      hWndList.Add CStr(hwnd)
      TitelList.Add Title
      
      'Elternfenster ermitteln
      Parent = hwnd
      Do
        Parent = GetParent(Parent)
      Loop Until Parent = 0
      
      'Task Id ermitteln
      Result = GetWindowThreadProcessId(hwnd, Task)
      ProcIDList.Add Task
    End If
    Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' GetWindowInfo

Private Sub Ini()
 On Error GoTo fehler
 Set hWndList = New Collection
 Set TitelList = New Collection
 Set ProcIDList = New Collection
 Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Ini/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Ini

Public Sub EnumWindows()
  Dim hwnd&
  On Error GoTo fehler
    Call Ini
    'Auch der Desktop ist ein Fenster
    hwnd = GetDesktopWindow
    Call GetWindowInfo(hwnd)
    
    'Einstieg
'    hWnd = GetWindow(hWnd, GW_HWNDFIRST)
    
    'Alle vorhandenen Fenster abklappern
      
    hwnd = GetWindow(hwnd, 5)
    
    Do
      Call GetWindowInfo(hwnd)
      hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop Until hwnd = 0
    Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in EnumWindows/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' EnumWindows

Public Function WartAufProzeß%(id&)
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
On Error GoTo fehler
'Dim collProcesses As New Collection
'Dim ProcID AS Long
  

    ' WINDOWS NT / 2000 / XP / 2003 / Vista
    ' -------------------------------------
  
    Dim cb&, cbNeeded&
'    Dim RetVal AS Long
    Dim NumElements&
    Dim ProcessIDs&()
    Dim cbNeeded2 As Long
'    Dim NumElements2 AS Long
'    Dim Modules(1) AS Long
'    Dim ModuleName AS String
'    Dim LenName AS Long
    Dim i&
  
    cb = 8         ' "CountBytes": Größe des Arrays (in Bytes)
    cbNeeded = 9   ' cbNeeded muss initial größer als cb sein
  
    ' Schrittweise an die passende Größe des Prozess-ID-Arrays
    ' heran-arbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' in cbNeeded steht der übergebene Speicherplatz in Bytes.
    ' Da jedes Element des Arrays als Long aus 4 Bytes besteht,
    ' ermitteln wir die Anzahl der tatsächlich übergebenen
    ' Elemente durch entsprechende Division:
    NumElements = cbNeeded / 4
  
    ' Jede gefundene Prozess-ID des Arrays abarbeiten
    WartAufProzeß = 0
    For i = 1 To NumElements
      If ProcessIDs(i) = id Then
        WartAufProzeß = -1
        Exit Function
      End If
    Next i
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WartAufProzeß/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WartAufProzeß

Function WarteAufAlt(Titel$, Tmax#)
 Dim hwnd&, i&, zl&, Zieli&, T1#, T2#
 On Error Resume Next
 AppActivate Titel
 On Error GoTo fehler
 T1 = Now
 Zieli = -1
 Do
  Call EnumWindows
  hwnd = GetForegroundWindow
  zl = 0
  Do
   For i = 1 To hWndList.COUNT
    If hwnd = hWndList(i) Then
     Zieli = i
     Exit For
    End If
    FPos = -273
   Next i
   If Zieli > 0 Then
    If InStrB(LCase$(TitelList(Zieli)), LCase$(Titel)) <> 0 Then Exit Do
   End If
   zl = zl + 1
   hwnd = GetParent(hwnd)
   If hwnd = 0 Then Exit Do
  Loop
  If InStrB(LCase$(TitelList(Zieli)), LCase$(Titel)) <> 0 Then
   WarteAufAlt = -1
   Exit Do
  End If
  T2 = Now
  If (T2 - T1) * 60 * 24 * 60 > Tmax Then Exit Do
 Loop
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WarteAufAlt/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WarteAufAlt

Function SchauObDa(Titel$, Optional obdebug%)
  Const fDatei0$ = "c:\schauobdadebug.txt"
  Dim i&
  On Error GoTo fehler
  Call EnumWindows
  If obdebug Then Open fDatei0 For Output As #337
  For i = 1 To hWndList.COUNT
   If obdebug Then Print #337, TitelList(i)
   If InStrB(TitelList(i), Titel) <> 0 Then
    SchauObDa = True
    Exit For
   End If
  Next i
  If obdebug Then
   Close #337
   If SchauObDa = 0 Then
'    Shell ("notepad " & fDatei0)
'    SuSh "notepad " & fDatei0, 2, , 0, 1
'    rufauf "notepad", fDatei0, , , 0, 1
     zeigan fDatei0
   End If
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SchauObDa/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' Schauobda

Function WarteAuf(Titel$, Tmax#)
 Dim sysdrv$
 sysdrv = LCase$(Environ("systemdrive"))
 Dim fDatei0$
 fDatei0$ = sysdrv & "\p1"
 Dim hwnd&, i&, zl&, Zieli&, T1#, T2#, fDatei$, nr&
 On Error Resume Next
 AppActivate Titel
 On Error GoTo fehler
 T1 = Now
 Zieli = -1
 Do
  Call EnumWindows
  On Error Resume Next
  nr = 0
  fDatei = fDatei0 & nr
  Do
   Err.Clear
   On Error Resume Next
   Close #333
   On Error GoTo fehler
   Open fDatei & ".txt" For Output As #333
   nr = nr + 1
   fDatei = fDatei0 & nr
   If Err.Number = 0 Then Exit Do
  Loop
  
  On Error GoTo fehler
  For i = 1 To hWndList.COUNT
   Print #333, TitelList(i)
   If InStrB(LCase$(TitelList(i)), LCase$(Titel)) <> 0 Then
    WarteAuf = True
    Exit Do
   End If
  Next i
  Close #333
  T2 = Now
  If (T2 - T1) * 60 * 24 * 60 > Tmax Then
   Exit Do
  End If
 Loop
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WarteAuf/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WarteAuf

Function WarteAufNicht(Titel$, Tmax#)
 Dim hwnd&, i&, zl&, Zieli&, T1#, T2#
 On Error Resume Next
 AppActivate Titel
 On Error GoTo fehler
 T1 = Now
 Zieli = -1
 Do
  Call EnumWindows
'  Open "u:\liste15.txt" For Output AS #332
  zl = 0
  For i = 1 To hWndList.COUNT
'    Print #332, TitelList(i)
    If InStrB(LCase$(TitelList(i)), LCase$(Titel)) <> 0 Then
     zl = zl + 1
    End If
  Next i
'  Close #332
  DoEvents
  If zl = 0 Then
     WarteAufNicht = True
     Exit Do
  End If
  T2 = Now
  If (T2 - T1) * 60 * 24 * 60 > Tmax Then Exit Do
 Loop
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in WarteAufNicht/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WarteAufNicht

'ShellAndWaitforTermination, von http://www.vbaccelerator.com/home/VB/Code/Libraries/Shell_Projects/Shell_And_Wait_For_Completion/article.asp
Public Function ShellaW( _
        sShell As String, _
        Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbNormalFocus, _
        Optional ByRef sError As String, _
        Optional ByVal lTimeOut As Long = 2000000000 _
    ) As Boolean
Dim hProcess As Long
Dim lR As Long
Dim lTimeStart As Long
Dim bSuccess As Boolean

On Error GoTo ShellAndWaitForTerminationError
    ' This is v2 which is somewhat more reliable:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    If (hProcess = 0) Then
        sError = "This program could not determine whether the process started." & _
             "Please watch the program AND check it completes."
        ' Only fail IF there is an error - this can happen
        ' when the program completes too quickly.
    Else
        bSuccess = True
        lTimeStart = timeGetTime()
        Do
            ' Get the status of the process
            GetExitCodeProcess hProcess, lR
            ' Sleep during wait to ensure the other process gets
            ' processor slice:
            DoEvents: Sleep 100
            If (timeGetTime() - lTimeStart > lTimeOut) Then
                ' Too long!
                sError = "The process has timed out."
                lR = 0
                bSuccess = False
            End If
        Loop While lR = STILL_ACTIVE
    End If
    ShellaW = bSuccess
    Exit Function
ShellAndWaitForTerminationError:
    sError = Err.Description
    Exit Function
End Function ' ShellaW

'SuperShell
'alsAdm: 0 = nein, 1= ja, 2= falls mit 0 kein Ergebnis, 3 = beide immer
Public Function SuSh&(ByVal App$, Optional alsAdm&, Optional ByVal WorkDir$, Optional dwmillis& = 10000, _
    Optional ByVal start_size& = SW_HIDE, Optional obkill&, Optional ByVal Priority_Class& = NORMAL_PRIORITY_CLASS)
    If WorkDir = vbNullString Then WorkDir = Environ("userprofile")
    Dim pclass&
    Dim runde&
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    'Not used, but needed
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    Dim lRetValue&
    'Set the structure size
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sInfo.cb = Len(sInfo)
    'Set the flags
    sInfo.dwFlags = STARTF_USESHOWWINDOW
    'Set the window's startup position
    sInfo.wShowWindow = start_size
    'Set the priority class
    pclass = Priority_Class
    'Start the program
    For runde = 1 To 2
'     IF CreateProcess(vbNullString, IIf(alsAdm = 1 OR (alsAdm = 2 AND runde = 2), vverz & doAlsAd & acceu & ap1 & ap2 & " ", "") & App, _
                     sec1, sec2, False, pclass, 0&, WorkDir, sinfo, pinfo) THEN
     If (alsAdm = 1 And runde = 1) Or ((alsAdm = 2 Or alsAdm = 3) And runde = 2) Then
      App = vVerz & doalsAd & acceu & ap1 & ap2 & " " & App ' wird nur einmal aufgerufen
     End If
     If CreateProcess(vbNullString, App, ByVal sec1, ByVal sec2, 0&, pclass, ByVal 0&, WorkDir, sInfo, pInfo) Then
        'Wait
        If dwmillis <> 0 Then WaitForSingleObject pInfo.hProcess, dwmillis
        SuSh = True
     Else
        SuSh = False
     End If
     If obkill Then
      lRetValue = TerminateProcess(pInfo.hProcess, 0&)
      lRetValue = CloseHandle(pInfo.hThread)
      lRetValue = CloseHandle(pInfo.hProcess)
     End If
     If (alsAdm = 2 And SuSh) Or (alsAdm < 2) Then Exit For
    Next runde
End Function ' SuSh

#If ersatz Then
Public Function m1tuAufruf&(ByRef App$, Optional alsAdm&, Optional ByVal WorkDir$, Optional dwmillis& = 10000, Optional ByVal start_size& = 0, Optional ByVal Priority_Class& = NORMAL_PRIORITY_CLASS)
         Dim pInfo As PROCESS_INFORMATION
         Dim sInfo As STARTUPINFO
         Dim sNull As String
         Dim lSuccess As Long
         Dim lRetValue As Long
         If WorkDir = vbNullString Then WorkDir = Environ("userprofile")
         sInfo.cb = Len(sInfo)
         Dim runde&
         For runde = 1 To 2
         lSuccess = CreateProcess(sNull, _
                                 IIf(alsAdm = 1 Or (alsAdm = 2 And runde = 2), vVerz & doalsAd & acceu & ap1 & ap2 & " ", "") & App, _
                                 ByVal 0&, _
                                 ByVal 0&, _
                                 1&, _
                                 Priority_Class, _
                                 ByVal 0&, _
                                 WorkDir, _
                                 sInfo, _
                                 pInfo)

         If lSuccess Then
             WaitForSingleObject pInfo.hProcess, dwmillis
         m1tuAufruf = True
        Else
         m1tuAufruf = False
        End If
        lRetValue = TerminateProcess(pInfo.hProcess, 0&)
        lRetValue = CloseHandle(pInfo.hThread)
        lRetValue = CloseHandle(pInfo.hProcess)
        If m1tuAufruf Or (alsAdm < 2) Then Exit For
       Next runde
      End Function
#End If

' GetForegroundWindow
Function GFGW&()
 Dim hwnd&, i&, zl&
 On Error GoTo fehler
 Call EnumWindows
 hwnd = GetForegroundWindow
 zl = 0
 Do
  For i = 1 To hWndList.COUNT
   If hwnd = hWndList(i) Then
'    Debug.Print CStr(zl) + ": " + TitelList(i)
    Exit For
   End If
  Next i
  zl = zl + 1
  hwnd = GetParent(hwnd)
  If hwnd = 0 Then Exit Do
 Loop
 GFGW = hwnd
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GFGW/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GFGW

Public Sub SwitchTo(hwnd&)
 Dim ret&, wStyle& ' Window Style bits
 On Error GoTo fehler

' Get style bits for window
 wStyle = GetWindowLong(hwnd, GWL_STYLE)
' IF minimized do a restore
 If wStyle And WS_MINIMIZE Then
  ret = ShowWindow(hwnd, SW_RESTORE)
 End If
' Move window to top of z-order/activate; no move/resize
 ret = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
 Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SwitchTo/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SwitchTo
  
' ----------------------------- CODE ------------------------------
  
' integriert das z.Zt. nicht funktionierende StopFos
Public Function GetProcessCollection(Optional obkill%, Optional Exe$) As Collection
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
  Dim collProcesses As New Collection
  Dim ProcID As Long
  Dim Hdl&
  On Error GoTo fehler
  If (Not IsWindowsNT) Then
  
    ' WINDOWS 95 / 98 / Me
    ' --------------------
  
    Dim sName As String
    Dim hSnap As Long
    Dim pEntry As PROCESSENTRY32
  
    ' Einen Snapshot der Prozessinformationen erstellen
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = 0 Then
      Exit Function ' Pech gehabt
    End If
  
    pEntry.dwSize = Len(pEntry) ' Größe der Struktur zur Verfügung stellen
  
    ' Den ersten Prozess im Snapshot ermitteln
    ProcID = Process32First(hSnap, pEntry)
  
    ' Mittels Process32Next über alle weiteren Prozesse iterieren
    Do While (ProcID <> 0) ' Gibt es eine gültige Prozess-ID?
      sName = TrimNullChar(pEntry.szExeFile)  ' Rückgabestring stutzen
      collProcesses.Add sName & "|" & CStr(ProcID) ' Collection-Eintrag
          If obkill <> 0 Then
           If InStrB(LCase$(sName), LCase$(Exe)) <> 0 Then
            Select Case obkill
             Case 1
              Hdl = GetWinHandle(ProcID)
              If Hdl <> 0 Then
               ShowWindow Hdl, SW_RESTORE
               SetForegroundWindow Hdl
               PostMessage Hdl, WM_CLOSE, 0&, 0&
               Exit Function
              End If
             Case 2
              If KillProcessByPID(ProcID) <> 0 Then Exit Function
            End Select
           End If
          End If
      ProcID = Process32Next(hSnap, pEntry)   ' Nächste PID des Snapshots
    Loop
  
    ' Handle zum Snapshot freigeben
    CloseHandle hSnap
  
  Else
    ' WINDOWS NT / 2000 / XP / 2003 / Vista
    ' -------------------------------------
    Dim cb As Long
    Dim cbNeeded As Long
    Dim RetVal As Long
    Dim NumElements As Long
    Dim ProcessIDs() As Long
    Dim cbNeeded2 As Long
    Dim NumElements2 As Long
    Dim Modules(1) As Long
    Dim ModuleName As String
    Dim LenName As Long
    Dim hProcess As Long
    Dim i As Long
  
    cb = 8         ' "CountBytes": Größe des Arrays (in Bytes)
    cbNeeded = 9   ' cbNeeded muss initial größer als cb sein
  
    ' Schrittweise an die passende Größe des Prozess-ID-Arrays
    ' heran-arbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' in cbNeeded steht der übergebene Speicherplatz in Bytes.
    ' Da jedes Element des Arrays als Long aus 4 Bytes besteht,
    ' ermitteln wir die Anzahl der tatsächlich übergebenen
    ' Elemente durch entsprechende Division:
    NumElements = cbNeeded / 4
  
    ' Jede gefundene Prozess-ID des Arrays abarbeiten
    For i = 1 To NumElements
      ' Versuchen, den Prozess zu öffnen und ein Handle zu erhalten
      hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                          Or PROCESS_VM_READ, _
                             0, ProcessIDs(i))
  
      If (hProcess <> 0) Then ' OpenProcess war erfolgreich
    
        ' EnumProcessModules gibt die dem Prozess angehörenden
        ' Module in einem Array zurück.
        RetVal = EnumProcessModules(hProcess, Modules(1), 1, cbNeeded2)
  
        If (RetVal <> 0) Then ' EnumProcessModules war erfolgreich
          ModuleName = Space$(MAX_PATH) ' Speicher reservieren
          ' Den Pfadnamen für das erste gefundene Modul bestimmen
          LenName = GetModuleFileNameEx(hProcess, Modules(1), ModuleName, Len(ModuleName))
          ' Den gefundenen Pfad und die Prozess-ID unserer
          ' ProcessCollection hinzufügen (Trennzeichen "|")
          collProcesses.Add Left$(ModuleName, LenName) & "|" & CStr(ProcessIDs(i))

          If obkill <> 0 Then
'           Debug.Print ModuleName
           If InStrB(LCase$(ModuleName), LCase$(Exe)) <> 0 Then
            Select Case obkill
             Case 1, -1
              Hdl = GetWinHandle(ProcessIDs(i))
              If Hdl <> 0 Then
               ShowWindow Hdl, SW_RESTORE
               SetForegroundWindow Hdl
               PostMessage Hdl, WM_CLOSE, 0&, 0&
               CloseHandle hProcess ' Offenes Handle schließen
               Exit Function
              End If
             Case 2
              If KillProcessByPID(ProcessIDs(i)) <> 0 Then
               Exit Function
              Else
'               Call Shell(vverz & "pcwkill.exe " & exe, vbMaximizedFocus)
'               SuSh vVerz & "pcwkill.exe " & Exe, 2, , 0, 3
               rufauf vVerz & "pcwkill", Exe, 2
               CloseHandle hProcess ' Offenes Handle schließen
               Exit Function
              End If
            End Select
           End If
          End If
        End If
      End If
      CloseHandle hProcess ' Offenes Handle schließen
    Next i
  End If
  ' Zusammengestellte Collection übergeben
  Set GetProcessCollection = collProcesses
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetProcessCollection /" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetProcessCollection(Optional obKill%, Optional exe$) AS Collection
   
Public Function ProcessName(ByVal CollectionString As String) As String
' Extrahiert aus einem String der Collection den Prozessnamen.
  Dim Pos1&
  On Error GoTo fehler
  ' Trenner suchen
  Pos1 = InStr(CollectionString, "|")
  ' Wenn Trenner vorhanden, Eintrag zurückgeben (sonst vbNullString)
  If (Pos1 > 0) Then
    ProcessName = Left$(CollectionString, Pos1 - 1)
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProcessName/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ProcessName(ByVal CollectionString AS String) AS String
  
  
Public Function ProcessHandle(ByVal CollectionString As String) As Long
' Extrahiert aus einem String der Collection die Prozess-ID.
  Dim Pos1 As Long
  On Error GoTo fehler
  ' Trenner suchen
  Pos1 = InStr(CollectionString, "|")
  ' Wenn Trenner vorhanden, Handle zurückgeben (sonst 0)
  If (Pos1 > 0) And (Len(CollectionString) > Pos1) Then
    ProcessHandle = CLng(Mid$(CollectionString, Pos1 + 1))
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProcessHandle/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ProcessHandle(ByVal CollectionString AS String) AS Long
  
  
Public Function KillProcessByPID(ByVal pid As Long) As Boolean
' Versucht auf Basis einer Prozess-ID, den zugehörigen
' Prozess zu terminieren. Im Erfolgsfall wird True zurückgegeben.
  Dim hProcess As Long, tpid&, p&, nRet
  On Error GoTo fehler
  ' Öffnen des Prozesses über seine Prozess-ID
  hProcess = OpenProcess(PROCESS_TERMINATE, False, pid)
  If hProcess = 0 Then
   If True Then
' hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    hProcess = OpenProcess(0, False, pid)
    GetExitCodeProcess hProcess, nRet
    Call TerminateProcess(hProcess, nRet)
    Call CloseHandle(hProcess)
   ElseIf True Then
    hProcess = OpenProcess(SYNCHRONIZE, False, pid)
    Dim jh&, n As SECURITY_ATTRIBUTES
    jh = CreateJobObject(n, "kill")
    GetWindowThreadProcessId pid, tpid
    p = OpenProcess(2035711, 0, pid)
    p = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    AssignProcessToJobObject jh, p
    TerminateJobObject jh, 0
   End If
  End If
  ' Gibt es ein Handle, wird der Prozess darüber abgeschossen
  If (hProcess <> 0) Or True Then
    KillProcessByPID = TerminateProcess(hProcess, 1&) <> 0
    CloseHandle hProcess
  Else
   hProcess = OpenProcess(PROCESS_VM_READ, False, pid)
   hProcess = OpenProcess(&H1, False, pid)
'   Stop
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in KillProcessByPID/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' KillProcessByPID(ByVal pid AS Long) AS Boolean
  
' ----------------- PRIVATE FUNKTIONEN ----------------------------

Private Function TrimNullChar(ByVal s As String) As String
' Kürzt einen String s bis zum Zeichen vor einem vbNullChar
  Dim Pos1 As Long
  On Error GoTo fehler
  ' vbNullChar = Chr$(0) im String suchen
  Pos1 = InStr(s, vbNullChar)
  ' Falls vorhanden, den String entsprechend kürzen
  If (Pos1 > 0) Then
    TrimNullChar = Left$(s, Pos1 - 1)
  Else
    TrimNullChar = s
  End If
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TrimNullChar/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TrimNullChar(ByVal s AS String) AS String
  
Private Function IsWindowsNT() As Boolean
' Gibt True für Windows NT (und 2000, XP, 2003, Vista) zurück
Dim OSInfo As OSVERSIONINFO
  On Error GoTo fehler
  With OSInfo
    .dwOSVersionInfoSize = Len(OSInfo)  ' Angabe der Größe dieser Struktur
    .szCSDVersion = Space$(128)         ' Speicherreservierung für Angabe des Service Packs
    GetVersionEx OSInfo                 ' OS-Informationen ermitteln
    IsWindowsNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT) ' für Windows NT
  End With
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in IsWindowsNT/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' IsWindowsNT() AS Boolean
  
Function ProcIDFromWnd(ByVal hwnd As Long) As Long
   Dim idProc As Long
   On Error GoTo fehler
   ' Get PID for this HWnd
   GetWindowThreadProcessId hwnd, idProc

   ' Return PID
   ProcIDFromWnd = idProc
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ProcIDFromWnd/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ProcIDFromWnd(ByVal hwnd AS Long) AS Long
 
Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
   On Error GoTo fehler
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)

   ' Loop until you find a match OR there are no more window handles:
   Do Until tempHwnd = 0
      ' Check IF no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If

      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWinHandle/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetWinHandle(hInstance AS Long) AS Long

Sub schließ_direkt(lWindowTitle$)
  Dim lHwnd As Long
  On Error GoTo fehler
'Dim lWindowTitle AS String
'    lWindowTitle = "FastObjects Server 9.5"
    lHwnd = FindWindow(vbNullString, lWindowTitle)
    If lHwnd = 0 Then
     lHwnd = FensterHandle(lWindowTitle)
    End If
    If lHwnd > 0 Then
        PostMessage lHwnd, WM_CLOSE, 0&, 0&
    End If
  Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in schließ_direkt/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' schließ_direkt

Function FensterHandle(Titel As String, Optional nr As Integer, Optional keinGroßKlein As Boolean) As Long
Dim hwnd&, hW2&, gefunden%
Dim ergstr As String * 255
Dim lenge As Integer, e1 As Long
Const consta As Integer = 2
On Error GoTo fehler
hwnd = GetForegroundWindow
Dim zaehler As Integer
zaehler = 1
Do While zaehler < 2000
 lenge = GetWindowTextLength(hwnd) + 1
 ergstr = vNS
 e1 = GetWindowText(hwnd, ergstr, lenge)
 If e1 > 0 Then
  If (Not keinGroßKlein And InStrB(ergstr, Titel) <> 0) Or (keinGroßKlein And InStrB(LCase$(ergstr), LCase$(Titel)) <> 0) Then
   nr = zaehler
   FensterHandle = hwnd
   Exit Function
  End If
 End If
 'Debug.Print LEFT(ergstr, 50), lenge, hwnd, e1
 hW2 = GetWindow(hwnd, consta)
 If hW2 = hwnd Then
    gefunden = True
    Exit Do
 End If
 hwnd = hW2
 zaehler = zaehler + 1
Loop
FensterHandle = 0
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FensterHandle/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' FensterHandle(Titel AS String, Optional Nr As Integer, Optional keinGroßKlein AS Boolean) AS Long
