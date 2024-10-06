Attribute VB_Name = "WartAufProzess"
Option Explicit
Private Declare Function GetParent& Lib "USER32" (ByVal hWnd&)
Private Declare Function GetForegroundWindow& Lib "USER32" ()
Private Declare Function GetWindowLong& Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal wIndx&)
Private Declare Function GetWindowTextLength& Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd&)
Private Declare Function GetWindowText& Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd&, ByVal lpString$, ByVal cch&)
Private Declare Function SetWindowPos& Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowThreadProcessId& Lib "USER32" (ByVal hWnd&, lpdwProcessId&)
Private Declare Function GetWindow& Lib "USER32" (ByVal hWnd&, ByVal wCmd&)
Private Declare Function EnumProcesses& Lib "psapi.dll" (ByRef lpidProcess&, ByVal cb&, ByRef cbNeeded&)
Private Declare Function ShowWindow& Lib "user32.dll" (ByVal hWnd&, ByVal nCmdShow&)
Private Declare Function GetDesktopWindow& Lib "USER32" ()
' Folgende Deklarationen alle für StopFos
Public Declare Function TerminateProcess& Lib "kernel32" (ByVal hProcess&, ByVal uExitCode&)
Private Declare Function lstrlenW& Lib "kernel32.dll" (ByVal StrPtr&)
Private Declare Function FindWindow& Lib "USER32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function PostMessage& Lib "USER32" Alias "PostMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Sub SetForegroundWindow Lib "USER32" (ByVal hWnd&)
Private Declare Function GetModuleFileNameEx& Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
  ByVal hProcess&, _
  ByVal hModule&, _
  ByVal ModuleName$, _
  ByVal nSize&)
Private Declare Function CreateToolhelp32Snapshot& Lib "kernel32" (ByVal dwFlags&, ByVal th32ProcessID&)
Private Declare Function Process32First& Lib "kernel32" (ByVal hSnapshot&, ByRef lppe As PROCESSENTRY32)
Private Declare Function Process32Next& Lib "kernel32" (ByVal hSnapshot&, ByRef lppe As PROCESSENTRY32)
' PSAPI-Funktionen zur Prozessauflistung (Windows NT)
Public Declare Function CloseHandle& Lib "kernel32.dll" (ByVal handle&)
Private Declare Function OpenProcess& Lib "kernel32.dll" (ByVal dwDesiredAccess&, ByVal bInheritHandle&, ByVal dwProcId&)
Private Declare Function EnumProcessModules& Lib "psapi.dll" (ByVal hProcess&, ByRef lphModule&, ByVal cb&, ByRef cbNeeded&)
' GetVersionEx dient der Erkennung des Betriebssystems:
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO)
' Win32-API-Funktionen für Prozessmanagement
' bis hierher für Stopfos
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
Private Const SW_RESTORE = 9 ' Restore window
Private Const SWP_SHOWWINDOW = &H40 ' Make window visible/active
Private Const HWND_TOP = 0 ' Move to top of z-order
Private Const SWP_NOSIZE = &H1 ' Do not re-size window
Private Const SWP_NOMOVE = &H2 ' Do not reposition window

Private Const obAuchTitellose% = -1
Private Const obAuchUnsichtbare% = -1

Private Const MAX_PATH = 260
Public Const WM_CLOSE = &H10
Private Const PROCESS_QUERY_INFORMATION As Long = 1024
Private Const PROCESS_VM_READ           As Long = 16
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_ALL_ACCESS        As Long = STANDARD_RIGHTS_REQUIRED _
                                               Or SYNCHRONIZE Or &HFFF
Private Const TH32CS_SNAPPROCESS        As Long = &H2&

' Konstante für die Erkennung des Betriebssystems:
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
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

Public TitelList As Collection, ProcIDList As Collection, hWndList As Collection
Dim FPos&
Dim ErrDescription$


Private Sub GetWindowInfo(ByVal hWnd&)
  Dim Parent&, Task&, Result&, x&, style&, Title$
  On Error GoTo fehler
  
    'Darstellung des Fensters
    style = GetWindowLong(hWnd, GWL_STYLE)
    style = style And (WS_VISIBLE Or WS_BORDER)
            
    'Title des Fenster auslesen
    Result = GetWindowTextLength(hWnd) + 1
    Title = Space$(Result)
    Result = GetWindowText(hWnd, Title, Result)
    Title = Left(Title, Len(Title) - 1)
    
    'In Abhängigkeit der Optionen die Ausgabe erstellen
    If (Title <> vNS Or obAuchTitellose) And _
       (style = (WS_VISIBLE Or WS_BORDER) Or obAuchUnsichtbare) Then
       
      hWndList.Add CStr(hWnd)
      TitelList.Add Title
      
      'Elternfenster ermitteln
      Parent = hWnd
      Do
        Parent = GetParent(Parent)
      Loop Until Parent = 0
      
      'Task Id ermitteln
      Result = GetWindowThreadProcessId(hWnd, Task)
      ProcIDList.Add Task
    End If
    Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
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
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Ini


Public Sub EnumWindows()
  Dim hWnd&
  On Error GoTo fehler
    Call Ini
    'Auch der Desktop ist ein Fenster
    hWnd = GetDesktopWindow
    Call GetWindowInfo(hWnd)
    
    'Einstieg
'    hWnd = GetWindow(hWnd, GW_HWNDFIRST)
    
    'Alle vorhandenen Fenster abklappern
      
    hWnd = GetWindow(hWnd, 5)
    
    Do
      Call GetWindowInfo(hWnd)
      hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop Until hWnd = 0
    Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' EnumWindows

Public Function WartAufProzeß%(ID&)
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
On Error GoTo fehler
'Dim collProcesses As New Collection
'Dim ProcID As Long
  

    ' WINDOWS NT / 2000 / XP / 2003 / Vista
    ' -------------------------------------
  
    Dim cb&, cbNeeded&
'    Dim RetVal As Long
    Dim NumElements&
    Dim ProcessIDs&()
'    Dim cbNeeded2 As Long
'    Dim NumElements2 As Long
'    Dim Modules(1) As Long
'    Dim ModuleName As String
'    Dim LenName As Long
    Dim i&
  
    cb = 8         ' "CountBytes": Größe des Arrays (in Bytes)
    cbNeeded = 9   ' cbNeeded muss initial größer als cb sein
  
    ' Schrittweise an die passende Größe des Prozess-ID-Arrays
    ' heranarbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' In cbNeeded steht der übergebene Speicherplatz in Bytes.
    ' Da jedes Element des Arrays als Long aus 4 Bytes besteht,
    ' ermitteln wir die Anzahl der tatsächlich übergebenen
    ' Elemente durch entsprechende Division:
    NumElements = cbNeeded / 4
  
    ' Jede gefundene Prozess-ID des Arrays abarbeiten
    WartAufProzeß = 0
    For i = 1 To NumElements
      If ProcessIDs(i) = ID Then
        WartAufProzeß = -1
        Exit Function
      End If
    Next i
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WartAufProzeß

Function WarteAuf(Titel$, TMax#)
 Dim hWnd&, i&, zl&, Zieli&, T1#, T2#
 On Error Resume Next
 AppActivate Titel
 On Error GoTo fehler
 T1 = Now
 Zieli = -1
 Do
  Call EnumWindows
  hWnd = GetForegroundWindow
  zl = 0
  Do
   For i = 1 To hWndList.Count
    If hWnd = hWndList(i) Then
     Zieli = i
     Exit For
    End If
    FPos = -273
   Next i
   If Zieli > 0 Then
    If InStrB(LCase(TitelList(Zieli)), LCase(Titel)) <> 0 Then Exit Do
   End If
   zl = zl + 1
   hWnd = GetParent(hWnd)
   If hWnd = 0 Then Exit Do
  Loop
  If InStrB(LCase(TitelList(Zieli)), LCase(Titel)) <> 0 Then
   WarteAuf = -1
   Exit Do
  End If
  T2 = Now
  If (T2 - T1) * 60 * 24 * 60 > TMax Then Exit Do
 Loop
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' WarteAuf
' GetForegroundWindow
Function GFGW2&()
 Dim hWnd&, i&, zl&
 Call EnumWindows
 hWnd = GetForegroundWindow
 zl = 0
 Do
  For i = 1 To hWndList.Count
   If hWnd = hWndList(i) Then
    Debug.Print CStr(zl) + ": " + TitelList(i)
    Exit For
   End If
  Next i
  zl = zl + 1
  hWnd = GetParent(hWnd)
  If hWnd = 0 Then Exit Do
 Loop
 GFGW2 = hWnd
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GFGW2

Public Sub SwitchTo2(hWnd&)
 Dim ret&, wStyle& ' Window Style bits
 On Error GoTo fehler

' Get style bits for window
 wStyle = GetWindowLong(hWnd, GWL_STYLE)
' If minimized do a restore
 If wStyle And WS_MINIMIZE Then
  ret = ShowWindow(hWnd, SW_RESTORE)
 End If
' Move window to top of z-order/activate; no move/resize
 ret = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
 Exit Sub
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Switchto/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' SwitchTo2
Private Function IsWindowsNT() As Boolean
' Gibt True für Windows NT (und 2000, XP, 2003, Vista) zurück
Dim OSInfo As OSVERSIONINFO
  
  With OSInfo
    .dwOSVersionInfoSize = Len(OSInfo)  ' Angabe der Größe dieser Struktur
    .szCSDVersion = Space$(128)         ' Speicherreservierung für Angabe des Service Packs
    GetVersionEx OSInfo                 ' OS-Informationen ermitteln
    IsWindowsNT = (.dwPlatformId = VER_PLATFORM_WIN32_NT) ' für Windows NT
  End With
  
End Function


Public Function StopFOS(exe$) As Collection
' Ermittelt die abfragbaren laufenden Prozesse des lokalen
' Rechners. Jeder gefundene Prozess wird mit seiner ID
' als String in einem Element der Rückgabe-Collection
' gespeichert im Format "Prozessname|Prozess-ID".
Dim collProcesses As New Collection
Dim ProcID As Long
On Error GoTo fehler:
exe = LCase(exe)
  
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
      sName = TrimNull(pEntry.szExeFile)  ' Rückgabestring stutzen
      collProcesses.Add sName & "|" & CStr(ProcID) ' Collection-Eintrag
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
    ' heranarbeiten. Dazu vergößern wir das Array großzügig immer
    ' weiter, bis der zur Verfügung gestellte Speicherplatz (cb)
    ' den genutzten (cbNeeded) überschreitet:
    Do While cb <= cbNeeded ' Alle Bytes wurden belegt -
                            ' es könnten also noch mehr sein
      cb = cb * 2                      ' Speicherplatz verdoppeln
      ReDim ProcessIDs(cb / 4) As Long ' Long = 4 Bytes
      EnumProcesses ProcessIDs(1), cb, cbNeeded ' Array abholen
    Loop
  
    ' In cbNeeded steht der übergebene Speicherplatz in Bytes.
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
        RetVal = EnumProcessModules(hProcess, Modules(1), _
                                    1, cbNeeded2)
  
        If (RetVal <> 0) Then ' EnumProcessModules war erfolgreich
          ModuleName = Space$(MAX_PATH) ' Speicher reservieren
          ' Den Pfadnamen für das erste gefundene Modul bestimmen
          LenName = GetModuleFileNameEx(hProcess, Modules(1), _
                                        ModuleName, Len(ModuleName))
'          Call GetWindowText(ProcessIDs(i), ModuleName, MAX_PATH) ' geht nicht, wenn in der Taskleiste
          If InStrB(LCase(ModuleName), exe) <> 0 Then '"FastObject") > 0 Then
'           Dim erg%
'           erg = KillProcessByPID(ProcessIDs(i))
           Dim hdl&
           hdl = GetWinHandle(ProcessIDs(i))
           If hdl <> 0 Then
            ShowWindow hdl, SW_RESTORE
            SetForegroundWindow hdl
            PostMessage hdl, WM_CLOSE, 0&, 0&
           End If
          Else
          ' Den gefundenen Pfad und die Prozess-ID unserer
          ' ProcessCollection hinzufügen (Trennzeichen "|")
          collProcesses.Add Left(ModuleName, LenName) & "|" & _
                            CStr(ProcessIDs(i))
          End If
        End If
  
      End If
  
      CloseHandle hProcess ' Offenes Handle schließen
  
    Next i
  
  End If
  
  ' Zusammengestellte Collection übergeben
  Set StopFOS = collProcesses
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function
Public Function KillProcessByPID(ByVal pid As Long) As Boolean
' Versucht auf Basis einer Prozess-ID, den zugehörigen
' Prozess zu terminieren. Im Erfolgsfall wird True zurückgegeben.
Dim hProcess As Long
  
  ' Öffnen des Prozesses über seine Prozess-ID
  hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
  
  ' Gibt es ein Handle, wird der Prozess darüber abgeschossen
  If (hProcess <> 0) Then
    PostMessage hProcess, WM_CLOSE, 0&, 0&
    KillProcessByPID = TerminateProcess(hProcess, 1&) <> 0
    CloseHandle hProcess
  End If
  
End Function
Function ProcIDFromWnd(ByVal hWnd As Long) As Long
   Dim idProc As Long

   ' Get PID for this HWnd
   GetWindowThreadProcessId hWnd, idProc

   ' Return PID
   ProcIDFromWnd = idProc
End Function

Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long

   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vNS, vNS)

   ' Loop until you find a match Or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
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
End Function
Private Function TrimNull$(startstr$)
   On Error GoTo fehler
   TrimNull = Left(startstr, lstrlenW(StrPtr(startstr)))
   Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetWindowInfo/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Sub schließ_direkt(lWindowTitle$)
Dim lHwnd As Long
'Dim lWindowTitle As String
'    lWindowTitle = "FastObjects Server 9.5"
    lHwnd = FindWindow(vNS, lWindowTitle)
    If lHwnd = 0 Then
     lHwnd = FensterHandle(lWindowTitle)
    End If
    If lHwnd > 0 Then
        PostMessage lHwnd, WM_CLOSE, 0&, 0&
    End If
End Sub

Function FensterHandle(Titel As String, Optional Nr As Integer, Optional keinGroßKlein As Boolean) As Long
Dim hWnd&, hW2&, gefunden%
Dim ergstr As String * 255
Dim lenge As Integer, e1 As Long
Const consta As Integer = 2
On Error GoTo fehler
hWnd = GetForegroundWindow
Dim zaehler As Integer
zaehler = 1
Do While zaehler < 2000
 lenge = GetWindowTextLength(hWnd) + 1
 ergstr = vNS
 e1 = GetWindowText(hWnd, ergstr, lenge)
 If e1 > 0 Then
  If (Not keinGroßKlein And InStrB(ergstr, Titel) <> 0) Or (keinGroßKlein And InStrB(LCase(ergstr), LCase(Titel)) <> 0) Then
   Nr = zaehler
   FensterHandle = hWnd
   Exit Function
  End If
 End If
 'Debug.Print left(ergstr, 50), lenge, hwnd, e1
 hW2 = GetWindow(hWnd, consta)
 If hW2 = hWnd Then
    gefunden = True
    Exit Do
 End If
 hWnd = hW2
 zaehler = zaehler + 1
Loop
FensterHandle = 0
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in FensterHandle/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' FensterHandle(Titel As String, Optional Nr As Integer, Optional keinGroßKlein As Boolean) As Long

