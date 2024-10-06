Attribute VB_Name = "SyncshellModul"
Option Explicit
Public Const NORMAL_PRIORITY_CLASS      As Long = &H20&
Public Const INFINITE                   As Long = -1&
Public Const STATUS_WAIT_0              As Long = &H0
Public Const WAIT_OBJECT_0              As Long = STATUS_WAIT_0
Private Declare Function InputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Type STARTUPINFO
cb              As Long
lpReserved      As Long
lpDesktop       As Long
lpTitle         As Long
dwX             As Long
dwY             As Long
dwXSize         As Long
dwYSize         As Long
dwXCountChars   As Long
dwYCountChars   As Long
dwFillAttribute As Long
dwFlags         As Long
wShowWindow     As Integer
cbReserved2     As Integer
lpReserved2     As Long
hStdInput       As Long
hStdOutput      As Long
hStdError       As Long
End Type
Public Type PROCESS_INFORMATION
hProcess    As Long
hThread     As Long
dwProcessID As Long
dwThreadID  As Long
End Type
Public Function Syncshell(CommandLine As String, Optional Timeout As Long, Optional WaitForInputIdle As Boolean, Optional Hide As Boolean = False) As Boolean
    Dim hProcess As Long
    Dim ret As Long
    Dim nMilliseconds As Long
    If Timeout > 0 Then
    nMilliseconds = Timeout
    Else
    nMilliseconds = INFINITE
    End If
    hProcess = StartProcess(CommandLine, Hide)
    If WaitForInputIdle Then        ' Warten, bis die eingeschlossene Anwendung        ' mit dem Erstellen ihrer Schnittstelle fertig ist:
    ret = InputIdle(hProcess, nMilliseconds)
    Else        ' Warten, bis die eingeschlossene Anwendung beendet ist:
    ret = WaitForSingleObject(hProcess, nMilliseconds)
    End If
    CloseHandle hProcess       ' "True" zurückgeben, wenn die Anwendung fertig ist.    ' Andernfalls Zeitüberschreitung oder Fehler.
    Syncshell = (ret = WAIT_OBJECT_0)
End Function
Public Function StartProcess(CommandLine As String, Optional Hide As Boolean = False) As Long
Const STARTF_USESHOWWINDOW As Long = &H1
Const SW_HIDE As Long = 0
Dim proc As PROCESS_INFORMATION
Dim Start As STARTUPINFO    ' STARTUPINFO-Struktur initialisieren:
Start.cb = Len(Start)
If Hide Then
Start.dwFlags = STARTF_USESHOWWINDOW
Start.wShowWindow = SW_HIDE
End If    ' Eingeschlossene Anwendung starten:
CreateProcessA 0&, CommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc
StartProcess = proc.hProcess
End Function
