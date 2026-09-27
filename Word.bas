Attribute VB_Name = "WinWord"
Option Explicit
Public Wapp  As Object ' AS Word.Application '
Dim WordWasNotRunning As Boolean ' Flag For final word unload
Public Const wdLineStyleSingle% = 1
Public Const wdWindowStateMaximize% = 1
Public Const wdFindContinue% = 1
Public Const wdAlignTabLeft% = 0
Public Const wdTabLeaderDots% = 1
Public Const wdReplaceAll% = 2
Public Const wdWord9TableBehavior% = 1
Public Const wdAutoFitContent% = 1
Public Const wdColorRed% = 255
Public Const wdBorderLeft% = -2
Public Const wdLineWidth050pt% = 4
Public Const wdColorAutomatic& = -16777216
Public Const wdBorderRight% = -4
Public Const wdBorderTop% = -1
Public Const wdBorderBottom% = -3
Public Const wdBorderHorizontal% = -5
Public Const wdBorderVertical% = -6
Public Const wdLineStyleNone% = 0
Public Const wdBorderDiagonalDown% = -7
Public Const wdBorderDiagonalUp% = -8
Public Declare Function sndPlaySound32& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName$, ByVal uFlags&)

' rechnet Zentimeter in Punkte (1/72 Zoll) um
' Aufruf in: Formular.LaborIns1, Formular.LaborTabAnp, Formular.tuBriefStandalone
Function CentimetersToPoints#(cm#)
 On Error GoTo fehler
 CentimetersToPoints = cm * 28.34646
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CentimetersToPoints/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' CentimetersToPoints

' Hauptversionsnummer des laufenden Word (Wapp.Build)
' Aufruf in: Formular.LaborIns1, Formular.tuBriefStandalone
Function WappBuild&()
 Dim Spl$()
 Spl = Split(CStr(Wapp.Build), ".")
 WappBuild = CLng(Spl(0))
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WappBuild/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WappBuild

' wartet (aktiv) sek Sekunden
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Function WarteSekunden(sek#)
 Dim T1#, T2#
 T1 = Now
 Do
  T2 = Now
  If (T2 - T1) * 60 * 60 * 24 > sek Then Exit Do
 Loop
End Function

' spielt die Wave-Datei Pfad asynchron ab
' Aufruf in: AnBog.vCommandB_Click, Formular.do_Diagnosen_Reset, Formular.do_HAC, Formular.Epikrise, Formular.Piep, Formular.PiepKurz, Formular.snie,
'   Formular.tuBriefStandalone
Public Function Sound(Pfad$)
 On Error GoTo fehler
 Call sndPlaySound32(Pfad, 1)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Sound/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

' aufgerufen in: do_anzeigen_click, do_PhotoImpact_Click, cmdPreview_Click, tuBriefStandalone, testWied
' setzt Wapp auf das laufende oder ein neu gestartetes Word
' 27.9.26: Aufruf in: Formular.cmdPreview_Click, Formular.do_anzeigen_click, Formular.do_PhotoImpact_Click, Formular.GetVorDat, Formular.testWied,
'   Formular.tuBriefStandalone, Formular.WordDateiOeffnen, Formular.WordDateiSchließen
Public Sub GetWord()
 On Error GoTo fehler
  Set Wapp = getAppl("OpusApp", "Word.Application")
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetWord/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' GetWord()

' in GetWord
' liefert die laufende Instanz der Anwendung ObjName (z.B. Word.Application) oder startet eine neue;
' merkt sich in WordWasNotRunning, ob sie schon lief
' 27.9.26: Aufruf in: WinWord.GetWord
Public Function getAppl(className, ObjName) As Object 'Word.Application

' Test to see IF there is a copy of Micr
' osoft Word already running.
'on error resume next' Defer error trapping.
' Getobject FUNCTION called without the
' first argument returns a
' reference to an instance of the applic
' ation. IF the application isn't
' running, an error occurs.
Dim FZahl%
On Error Resume Next
vonvorne:
Set getAppl = GetObject(, ObjName)
If Err.Number <> 0 Then
' syscmd 4, "getApp1, Fehler: " & Err.Number & ":" & Err.Description
 syscmd 4, "Muss 'word' noch aufrufen"
 WordWasNotRunning = True
Else
 WordWasNotRunning = False
End If
Err.Clear ' Clear Err object in Case Error occurred.
' Check for Microsoft Word. IF Microsoft
' Word is running,
' enter it INTO the Running Object table
' .
On Error GoTo fehler

If WordWasNotRunning Then
'Set the object variable to start a new
' instance of Word.
neu:
Select Case ObjName
 Case "Word.Application"
  Set getAppl = CreateObject(ObjName) 'wobj ' New Word.Application
 Case Else
End Select
End If
' Show Microsoft Word through its Applic
' ation property. THEN
' show the actual window containing the
' file USING the Windows
' collection of the MyWord object refere
' nce.
On Error Resume Next
'getAppl.Visible = True
If Err.Number <> 0 And FZahl < 10 Then
 FZahl = FZahl + 1
 GoTo vonvorne
End If
getAppl.Application.WindowState = wdWindowStateMaximize
Select Case Err.Number
 Case 0
 Case 5825: GoTo neu
End Select
Exit Function
On Error GoTo fehler
Screen.MousePointer = 0 ' vbDefault
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getAppl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getAppl

'Demo of how to call the above sub


