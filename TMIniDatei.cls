VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMIniDatei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Local.ini und Global.ini"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"TMIniDatei"
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
#Const obDebug = 0
Private mvLoc As TMIniSectionColl, mvLocAlt As TMIniSectionColl 'lokale Kopie
Private mvGlob As TMIniSectionColl, mvGlobAlt As TMIniSectionColl 'lokale Kopie
Public TMVerz$ ' TurbomedVerzeichnis mit Programm
Public LokalTurbomed$ ' TurbomedVerzeichnis ohne Programm
Public LocalIni$
Public GlobalIniS$
Public PraxDBV$

' FPos muß noch woanders definiert werden. klassenvariable habe ich nicht gefunden

' Klasse TMIniDatei: liest und schreibt die Turbomed-Einstellungsdateien Local.ini und Global.ini
' Class_Terminate: (Zurückschreiben beim Beenden ist abgeschaltet)
' Aufruf in: Ereignisprozedur
Private Sub Class_Terminate()
 On Error GoTo fehler
' Call Schreiben
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Terminate/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Terminate

' schreibt Local.ini und Global.ini von Turbomed zurück, soweit sie sich gegenüber dem Ursprungszustand geändert haben
' Aufruf in: TMIniDatei.Sichern
Private Sub Schreiben()
 On Error GoTo fehler
 Call iniDateiSchreib(LocalIni, mvLoc, mvLocAlt, 4)
 Call iniDateiSchreib(GlobalIniS, mvGlob, mvGlobAlt, 4)
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Schreiben/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Schreiben

' schreibt geänderte Einstellungen zurück und merkt sich den neuen Stand als Ursprungszustand
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Sub Sichern()
 On Error GoTo fehler
 Call Schreiben
 Call UrsprungsZustand
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Sichern/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Sichern

' setzt die Eigenschaft Nam im Abschnitt Sec (erst in Local.ini, sonst in Global.ini) auf Wert;
' gibt den bisherigen Inhalt bzw. eine Fehlermeldung zurück
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Function SetProp$(Sec$, Nam$, Wert$)
 Dim AltProp$
 On Error GoTo fehler
 Call ergSec(Sec)
 On Error Resume Next
 SetProp = mvLoc.Item(Sec).UnterEig(Nam).Inhalt
 If Err.Number = 0 Then
  mvLoc.Item(Sec).UnterEig(Nam).Inhalt = Wert
 ElseIf Err.Number = 91 Then
  Err.Clear
  SetProp = mvGlob.Item(Sec).UnterEig(Nam).Inhalt
  If Err.Number = 0 Then
   mvGlob.Item(Sec).UnterEig(Nam).Inhalt = Wert
  ElseIf Err.Number = 91 Then
   SetProp = "Fehler: Eigenschaft '" + Sec + Nam + "' nicht gefunden!"
'   Debug.Print SetProp
  End If
 End If
' ON Error GoTo fehler
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetProp/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetProp

' ergänzt den Abschnittsnamen um "TurboMed Grundeinstellungen/" vorne und "/" hinten
' Aufruf in: TMIniDatei.GetProp, TMIniDatei.SetProp
Function ergSec(Sec$)
 On Error GoTo fehler
 If InStr(Sec, "TurboMed Grundeinstellungen/") <> 1 Then
  Sec = "TurboMed Grundeinstellungen/" + Sec
 End If
 If Right$(Sec, 1) <> "/" Then Sec = Sec + "/"
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ErgSec/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ergSec

' liest die Eigenschaft Nam im Abschnitt Sec (erst aus Local.ini, sonst aus Global.ini)
' Aufruf in: Haupt.getDokPfad, Lese.DokumenteInDatenbank_Click
Public Function GetProp$(Sec$, Nam$)
 On Error GoTo fehler
 Call ergSec(Sec)
 On Error Resume Next
 GetProp = mvLoc.Item(Sec).UnterEig(Nam).Inhalt
 If Err.Number = 91 Then
  Err.Clear
  GetProp = mvGlob.Item(Sec).UnterEig(Nam).Inhalt
  If Err.Number = 91 Then
#If obDebug Then
   GetProp = "Fehler: Eigenschaft '" + Sec + "'/'" + Nam + "' nicht gefunden!"
#End If
  End If
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetProp/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetProp

' schreibt die Abschnitte mv in eine Temporärdatei und tauscht sie gegen Datei aus, wenn sie sich von mvalt unterscheiden;
' hält dabei ResZahl alte Fassungen (Datei & "alt1" ...) vor
' Aufruf in: TMIniDatei.Schreiben
Private Sub iniDateiSchreib(Datei, mv As TMIniSectionColl, mvalt As TMIniSectionColl, ResZahl%)
 Dim tmpDat$, i&, j&
 On Error GoTo fehler
 If Not mv.VergleicheMit(mvalt) Then
  tmpDat = REPLACE$(Datei, ".ini", "tmp.ini")
'  tmpDat = Environ("tmp") & "\Local.ini"
  On Error Resume Next
  Do
   Kill tmpDat
   If Err.Number <> 75 Then Exit Do
   tmpDat = tmpDat & "_"
  Loop
  On Error GoTo fehler
  Open tmpDat For Output As #11
   For i = 1 To mv.COUNT
    Print #11, "[" & mv.Item(i).Bezeichnung & "]"
    For j = 1 To mv.Item(i).UnterEig.COUNT
     Print #11, mv.Item(i).UnterEig.Item(j).Bezeichnung & "={" & mv.Item(i).UnterEig.Item(j).Inhalt & "}"
    Next j
   Next i
  Close #11
  DoEvents
  On Error Resume Next
  Kill Datei & "alt" & CStr(ResZahl)
  For i = ResZahl To 1 Step -1
   Name Datei & "alt" & CStr(i - 1) As Datei & "alt" & CStr(i)
  Next i
  On Error GoTo fehler
  Name Datei As Datei & "alt1"
  Name tmpDat As Datei
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in iniDateiSchreib/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' iniDateiSchreib

' liest eine Turbomed-ini-Datei (Zeilen "[Abschnitt]" und "Name={Inhalt}") in die Abschnittsliste mv
' Aufruf in: TMIniDatei.Class_Initialize
Private Sub DateiLes(Datei, mv As TMIniSectionColl)
 Dim Text$, sp$()
 Dim SPS$
 On Error GoTo fehler
 If Not FileExists(Datei) Then
  If FileExists(Datei + "alt") Then
   Name Datei + "alt" As Datei
  Else
   Exit Sub ' 8.1.08
  End If
 End If
 Open Datei For Input Access Read As #10
  Do While Not EOF(10)
   Line Input #10, Text
   If left$(Text, 1) = "[" Then
    Text = Mid$(Text, 2, Len(Text) - 2)
    Call mv.Add(Text)
   ElseIf LenB(Text) Then
    sp = Split(Text, "={")
    If UBound(sp) > 0 Then
     SPS = left$(sp(1), Len(sp(1)) - 1)
    Else
     SPS = vNS
    End If
    Call mv.UEAdd(sp(0), SPS)
   End If
  Loop
 Close #10
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "Datei: " + Datei + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DateiLes/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DateiLes

' merkt sich Kopien der gelesenen Einstellungen, um später Änderungen zu erkennen
' Aufruf in: TMIniDatei.Class_Initialize, TMIniDatei.Sichern
Private Sub UrsprungsZustand()
 On Error GoTo fehler
 Set mvLocAlt = mvLoc.Kopiere
 Set mvGlobAlt = mvGlob.Kopiere
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UrsprungsZustand/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' UrsprungsZustand

' ermittelt das Turbomed-Programmverzeichnis (Registry, sonst %Homedrive%\Turbomed\Programm), kopiert ggf. zu kleine
' Local.ini/Global.ini aus dem Turbomed-Stammverzeichnis und liest beide ein
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
 Dim AltEintr$
 Dim EintrZ
 On Error GoTo fehler
 Dim Eintr As New TMIniSectionColl
 Dim wsh As New WshShell ' %windir%\system32\wshom.ocx
 On Error Resume Next
 TMVerz = wsh.RegRead("HKEY_LOCAL_MACHINE\Software\TurboMed EDV GmbH\TurboMed\Current\RegisterPath")
 On Error GoTo fehler
 If TMVerz = vNS Then TMVerz = Environ("Homedrive") + "\Turbomed\Programm\"
 LokalTurbomed = left$(TMVerz, Len(TMVerz) - Len("\Programm")) ' c:\TurboMed
 LocalIni = TMVerz + "Local.ini" ' c:\Turbomed\Programm\Local.ini
 GlobalIniS = TMVerz + "Global.ini" ' c:\Turbomed\Programm\Global.ini
 Dim FLlocal&, FLglobal&
 If FileExists(LocalIni) Then
  FLlocal = FSO.GetFile(LocalIni).size ' 21.8.21 statt filelen
 Else
  FLlocal = 0
 End If
 If FileExists(GlobalIniS) Then
  FLglobal = FSO.GetFile(GlobalIniS).size
 Else
  FLglobal = 0
 End If
 If LenB(TMStammV) <> 0 Then
  VerzPrüf TMVerz
  If FLlocal < 15000 Then
   If FSO.FolderExists(TMStammV & "\Programm") Then
    KopDat TMStammV & "\Programm\Local.ini", LocalIni
   End If
  End If
  If FLglobal < 7000 Then
   If FSO.FolderExists(TMStammV & "\Programm") Then
    KopDat TMStammV & "\Programm\Global.ini", GlobalIniS
   End If
  End If
 End If
 Set mvLoc = New TMIniSectionColl
 Set mvGlob = New TMIniSectionColl
 Call DateiLes(LocalIni, mvLoc)
 Call DateiLes(GlobalIniS, mvGlob)
 Call UrsprungsZustand
 'AltEintr = SetProp("TurboMed Grundeinstellungen/Akustische Signale/Karteikarte/", "Medikamenten-Verordnung", "Lasix")
 'AltEintr = SetProp("TurboMed Grundeinstellungen/Datenbank/", "Uhrzeit Ende der Prüfung des Datensicherungsmodus", "23:59:59")
 'If mvGlob.VergleicheMit(mvGlobAlt) THEN
'  Debug.Print "gleich"
' Else
  'Debug.Print "verschieden"
 'END IF
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Initialize/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Initialize


' Abschnittsliste (setzen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Set Werte(ByVal vData As TMIniSectionColl)
 On Error GoTo fehler
 Set mvarWerte = vData
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Werte[PropertySet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Werte


' Abschnittsliste (lesen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get Werte() As TMIniSectionColl
 On Error GoTo fehler
 Set Werte = mvarWerte
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Werte[PropertyGet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Werte

