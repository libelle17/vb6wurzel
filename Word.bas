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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CentimetersToPoints/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' CentimetersToPoints

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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WappBuild/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WappBuild
Public Function WarteSekunden(sek#)
 Dim T1#, T2#
 T1 = Now
 Do
  T2 = Now
  If (T2 - T1) * 60 * 60 * 24 > sek Then Exit Do
 Loop
End Function

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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Sound/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetWord/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

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
 syscmd 4, "getApp1, Fehler: " & Err.Number & ":" & Err.Description
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

If WordWasNotRunning = True Then
'Set the object variable to start a new
' instance of Word.
neu:
Select Case ObjName
 Case "Word.Application"
  Set getAppl = CreateObject("Word.Application") 'wobj ' New Word.Application
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getAppl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getAppl
'Demo of how to call the above sub


