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
#Const obdebug = 0
Private mvLoc As TMIniSectionColl, mvLocAlt As TMIniSectionColl 'lokale Kopie
Private mvGlob As TMIniSectionColl, mvGlobAlt As TMIniSectionColl 'lokale Kopie
Public TMVerz$ ' TurbomedVerzeichnis mit Programm
Public LokalTurbomed$ ' TurbomedVerzeichnis ohne Programm
Public LocalIni$
Public GlobalIniS$
Public PraxDBV$

' FPos muß noch woanders definiert werden. klassenvariable habe ich nicht gefunden
Private Sub Class_Terminate()
 On Error GoTo fehler
' Call Schreiben
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Terminate/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Terminate

Private Sub Schreiben()
 On Error GoTo fehler
 Call iniDateiSchreib(LocalIni, mvLoc, mvLocAlt, 4)
 Call iniDateiSchreib(GlobalIniS, mvGlob, mvGlobAlt, 4)
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Schreiben/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Schreiben

Public Sub Sichern()
 On Error GoTo fehler
 Call Schreiben
 Call UrsprungsZustand
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Sichern/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Sichern

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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SetProp/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SetProp

Function ergSec(Sec$)
 On Error GoTo fehler
 If InStr(Sec, "TurboMed Grundeinstellungen/") <> 1 Then
  Sec = "TurboMed Grundeinstellungen/" + Sec
 End If
 If Right$(Sec, 1) <> "/" Then Sec = Sec + "/"
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ErgSec/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ergSec

Public Function GetProp$(Sec$, Nam$)
 On Error GoTo fehler
 Call ergSec(Sec)
 On Error Resume Next
 GetProp = mvLoc.Item(Sec).UnterEig(Nam).Inhalt
 If Err.Number = 91 Then
  Err.Clear
  GetProp = mvGlob.Item(Sec).UnterEig(Nam).Inhalt
  If Err.Number = 91 Then
#If obdebug Then
   GetProp = "Fehler: Eigenschaft '" + Sec + "'/'" + Nam + "' nicht gefunden!"
#End If
  End If
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetProp/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetProp
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in iniDateiSchreib/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' iniDateiSchreib

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
   If Left$(Text, 1) = "[" Then
    Text = Mid$(Text, 2, Len(Text) - 2)
    Call mv.Add(Text)
   ElseIf LenB(Text) Then
    sp = Split(Text, "={")
    If UBound(sp) > 0 Then
     SPS = Left$(sp(1), Len(sp(1)) - 1)
    Else
     SPS = vNS
    End If
    Call mv.UEAdd(sp(0), SPS)
   End If
  Loop
 Close #10
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "Datei: " + Datei + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DateiLes/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DateiLes

Private Sub UrsprungsZustand()
 On Error GoTo fehler
 Set mvLocAlt = mvLoc.Kopiere
 Set mvGlobAlt = mvGlob.Kopiere
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UrsprungsZustand/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' UrsprungsZustand

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
 LokalTurbomed = Left$(TMVerz, Len(TMVerz) - Len("\Programm")) ' c:\TurboMed
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Initialize/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Initialize


Public Property Set Werte(ByVal vData As TMIniSectionColl)
 On Error GoTo fehler
 Set mvarWerte = vData
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Werte[PropertySet]/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Werte


Public Property Get Werte() As TMIniSectionColl
 On Error GoTo fehler
 Set Werte = mvarWerte
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Werte[PropertyGet]/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Werte

