Attribute VB_Name = "FileTime"
Option Explicit
Private Declare Function OpenFile& Lib "kernel32" (ByVal lpFileName$, ByRef lpReOpenBuff As OFSTRUCT, ByVal wStyle&)
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject&)
Private Declare Function CreateFile Lib "kernel32.dll" _
  Alias "CreateFileA" ( _
  ByVal lpFileName As String, _
  ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, _
  lpSecurityAttributes As Any, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) As Long
  
Private Const GENERIC_READ = &H80000000 ' nur lesen
Private Const GENERIC_WRITE = &H40000000 ' nur schreiben
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
' CreateFile dwCreationDisposition Konstanten
Private Const CREATE_ALWAYS = 2
' erstellt eine neue Datei und überschreibt bereits vorhandene
Private Const CREATE_NEW = 1
' erstellt eine neue Datei nur dann, wenn sie noch nicht existiert
Private Const OPEN_ALWAYS = 4
' öffnet eine bereits vorhandene Datei und erstellt eine Datei, falls nicht vorhanden
Private Const OPEN_EXISTING = 3 ' öffnet eine bereits vorhandene Datei
Private Const TRUNCATE_EXISTING = 5
' öffnet eine bereits vorhandene Datei und löscht deren Inhalt.
 
' CreateFile dwFlagsAndAttributes
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20 ' Archiv-Datei
Private Const FILE_ATTRIBUTE_HIDDEN = &H2 ' Versteckt
Private Const FILE_ATTRIBUTE_NORMAL = &H80 ' Normal
Private Const FILE_ATTRIBUTE_READONLY = &H1 ' Schreibgeschützt
Private Const FILE_ATTRIBUTE_SYSTEM = &H4 ' Systemdatei
  
' AS Any-Deklaration von SetFileTime (fuer gewöhnlich
' FILETIME-Parameter, die aber zu NULL gesetzt werden können,
' um die entsprechende Information zu ignorieren
Private Declare Function SetFileTimeAPI& Lib "kernel32" Alias "SetFileTime" (ByVal hFile&, _
  ByRef lpCreationTime As Any, ByRef lpLastAccessTime As Any, ByRef lpLastWriteTime As Any)
' Gegenstück:
Private Declare Function GetFileTimeAPI _
  Lib "kernel32" Alias "GetFileTime" ( _
  ByVal hFile As Long, _
  ByRef lpCreationTime As Any, _
  ByRef lpLastAccessTime As Any, _
  ByRef lpLastWriteTime As Any _
  ) As Long
  
' Lokale Dateizeit in eine universale Dateizeit übersetzen
Private Declare Function LocalFileTimeToFileTime _
  Lib "kernel32" ( _
  ByRef lpLocalFileTime As FILETIME, _
  ByRef lpFileTime As FILETIME _
  ) As Long
' Gegenstück:
Private Declare Function FileTimeToLocalFileTime _
  Lib "kernel32" ( _
  ByRef lpFileTime As FILETIME, _
  ByRef lpLocalFileTime As FILETIME _
  ) As Long
  
' Eine SYSTEMTIME-Struktur in eine FILETIME übersetzen
Private Declare Function SystemTimeToFileTime _
  Lib "kernel32" ( _
  ByRef lpSystemTime As SYSTEMTIME, _
  ByRef lpFileTime As FILETIME _
  ) As Long
' Gegenstück:
Private Declare Function FileTimeToSystemTime _
  Lib "kernel32" ( _
  ByRef lpFileTime As FILETIME, _
  ByRef lpSystemTime As SYSTEMTIME _
) As Long
  
Private Const OF_READ = &H0
Private Const OF_READWRITE = &H2
Private Const OFS_MAXPATHNAME = 128
  
Private Type FILETIME
  dwLowDATETIME As Long
  dwHighDATETIME As Long
End Type
  
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
  
Private Type OFSTRUCT
  cBytes As Byte
  fFixedDisk As Byte
  nErrCode As Integer
  Reserved1 As Integer
  Reserved2 As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type
  
Public Enum FileTimeEnum
  mftCreationTime = 1
  mftLastAccessTime = 2
  mftlastwritetime = 4
End Enum
Dim FNr&
Const INVALID_FILE_VALUE$ = -1
Public Const HFILE_ERROR = &HFFFF

Public Function GetFileTime(ByVal Pfad As String, _
                            ByVal TimeToGet As FileTimeEnum _
                           ) As Date
' Ermittelt einen der drei Zeitstempel einer Datei/eines Verzeichnisses
' und gibt diesen als Visual Basic Date-Variable zurück.
Dim FTCreationTime As FILETIME, SysTime As SYSTEMTIME
Dim FTLastAccessTime As FILETIME
Dim ftLastWriteTime As FILETIME
Dim SELECTedTime As FILETIME
Dim OFS As OFSTRUCT, hFile As Long
  On Error GoTo fehler
  ' Versuchen, die betroffene Datei zu öffnen
'  hFile = OpenFile(Pfad, OFS, OF_READ)
'  Do
   hFile = CreateFile(Pfad, 0, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
'   IF hFile <> INVALID_FILE_VALUE THEN Exit Do
'   Debug.Print "GetFileTime: " & hFile
'  Loop
  If hFile = INVALID_FILE_VALUE Then Exit Function ' OpenFile ist gescheitert => Ausgang
  ' Ermitteln der Zeitstempel
  GetFileTimeAPI hFile, FTCreationTime, FTLastAccessTime, ftLastWriteTime
  CloseHandle hFile
  
  ' Gesuchten Zeitstempel auswählen
  Select Case TimeToGet
    Case mftCreationTime: SELECTedTime = FTCreationTime
    Case mftLastAccessTime: SELECTedTime = FTLastAccessTime
    Case mftlastwritetime: SELECTedTime = ftLastWriteTime
  End Select
  
  ' Umsetzung in lokale Systemzeit
  FileTimeToLocalFileTime SELECTedTime, SELECTedTime
  FileTimeToSystemTime SELECTedTime, SysTime
  
  ' Rückgabe als VB-Date
  With SysTime
    GetFileTime = _
      DateSerial(.wYear, .wMonth, .wDay) + _
      TimeSerial(.wHour, .wMinute, .wSecond)
  End With
  Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetFileTime/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetFileTime


Public Function SetFileTimeByDate(ByVal Pfad As String, _
                                  ByVal TimeToModify As FileTimeEnum, _
                                  ByVal DateToSet As Date)
 On Error GoTo fehler
' Setzt den Zeitstempel einer Datei unter Zuhilfenahme
' der ausführenden Funktion SetFileTime.
  
  SetFileTimeByDate = SetFileTime(Pfad, TimeToModify, Day(DateToSet), Month(DateToSet), Year(DateToSet), _
                                        Hour(DateToSet), Minute(DateToSet), Second(DateToSet))
  Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetFileTimeByDate/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetFileTimeByDate
  
  
Private Function SetFileTime(ByVal Pfad$, _
                             ByVal TimeToModify As FileTimeEnum, _
                             ByVal Tag%, _
                             ByVal Monat%, _
                             ByVal Jahr%, _
                             ByVal Stunde%, _
                             ByVal Minute%, _
                             ByVal Sekunde%) As Boolean ' True => Erfolg
  
Dim FT As FILETIME, ST As SYSTEMTIME, FTC As FILETIME, FTA As FILETIME, FTW As FILETIME
Dim OFS As OFSTRUCT, hFile&, RetVal&
  On Error GoTo fehler
  ' Dateizeiten (FILETIME) und Systemzeiten (SYSTEMTIME)
  ' unterscheiden sich im Format. Zunaechst wird eine
  ' SYSTEMTIME-Struktur mit den uebergebenen Parametern gefüllt,
  ' danach wird sie in eine FILETIME konvertiert. Da Dateizeiten
  ' GMT-orientiert geschrieben werden, ist danach noch eine Anpassung
  ' an die GMT-Zeitzone erforderlich.
  
  ' SYSTEMTIME-Struktur ausfüllen
  With ST
    .wYear = Jahr
    .wMonth = Monat
    .wDay = Tag
    .wHour = Stunde
    .wMinute = Minute
    .wSecond = Sekunde
    '.wMilliseconds = 0
  End With
  
  ' Lokale Systemzeit in lokale Dateizeit konvertieren
  RetVal = SystemTimeToFileTime(ST, FT)
  If RetVal = 0 Then Exit Function  ' Pech gehabt
  
  ' Lokale Dateizeit in GMT-Dateizeit konvertieren
  RetVal = LocalFileTimeToFileTime(FT, FT)
  If RetVal = 0 Then Exit Function  ' Pech gehabt
  
  ' Datei fuer Lese- und Schreibzugriff öffnen
'  hFile = OpenFile(Pfad, OFS, OF_READWRITE)
'  IF hFile = HFILE_ERROR THEN Exit FUNCTION ' Pech gehabt
  hFile = CreateFile(Pfad, GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
'  Debug.Print "SetFileTime: " & hFile
'  Loop
  If hFile = INVALID_FILE_VALUE Then Exit Function ' OpenFile ist gescheitert => Ausgang
  If 0 Then Call GetFileTimeAPI(hFile, FTC, FTA, FTW)
  ' Eine Datei hat 3 Dateizeiten:
  ' Erzeugung, Letzter Zugriff, Letzte Speicherung.
  ' Eine Zeit davon soll geändert werden, der Rest soll
  ' identisch bleiben. Dank "As Any"-Deklaration werden
  ' nicht zu ändernde Zeitparameter mit "ByVal 0&" (für
  ' NULL) bedient.
  
  If (TimeToModify And mftCreationTime) > 0 Then
    RetVal = SetFileTimeAPI(hFile, FT, ByVal 0&, ByVal 0&)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  If (TimeToModify And mftLastAccessTime) > 0 Then
    RetVal = SetFileTimeAPI(hFile, ByVal 0&, FT, ByVal 0&)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  If (TimeToModify And mftlastwritetime) > 0 Then
'    RetVal = SetFileTimeAPI(hFile, ByVal FTC, ByVal FTW, FT)
    RetVal = SetFileTimeAPI(hFile, ByVal 0&, ByVal 0&, FT)
    If RetVal = 0 Then CloseHandle hFile: Exit Function ' Pech gehabt
  End If
  
  ' Handle schließen
  CloseHandle hFile
  
  SetFileTime = True
  Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetFileTime/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetFileTime

