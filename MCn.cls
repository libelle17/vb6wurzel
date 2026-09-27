VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Cn As ADODB.Connection
Public obT% ' ob Transaktion offen
Public CNs$ ' connection string
Public tlev%
Public rAf&
Public ErrNr&
Public ErrDs$
Public FNr$

' Klasse MCn: ADO-Verbindung mit Merkfeldern für Verbindungszeichenfolge, Fehler und betroffene Sätze
' Class_Initialize: (ohne Aufgabe)
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
' Debug.Print "hier bin ich"
End Sub ' Class_Initialize()

' öffnet eine neue Verbindung mit der Verbindungszeichenfolge pCNs
' Aufruf in: Lese.mdiForm_Load
Public Sub init(pCNs$)
 On Error GoTo fehler
   Set Cn = Nothing
   Set Cn = New ADODB.Connection
   CNs = pCNs
   Cn.Open pCNs
   Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in MCn.init/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' init

' führt Befehl aus und merkt sich Fehlernummer, Fehlertext und Zahl der betroffenen Sätze (prAf)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Function führaus(Befehl$, Optional prAf&, Optional options& = -1) As ADODB.Recordset
 On Error GoTo fehler
 Set führaus = Cn.Execute(Befehl, prAf, options)
 ErrNr = Err.Number
 ErrDs = Err.Description
 rAf = prAf
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in MCn.führaus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' führaus

#If zutesten Then
' beginnt eine Transaktion (nur mit zutesten)
Public Sub Begin()
 On Error GoTo fehler
 Cn.Execute "START TRANSACTION"
 tlev = 1
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in MCn.Begin/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Begin

' schließt eine offene Transaktion ab
Public Sub Commit()
 On Error GoTo fehler
 If tlev Then
   Cn.Execute "COMMIT"
   tlev = 0
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "aufgefangener Fehler in MCn.Commit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Commit

' schließt beim Beenden die offene Transaktion ab und gibt die Verbindung frei
Private Sub Class_Terminate() ' wird auch bei Neuzuweisung ohne explitzites set mcn=nothing aufgerufen
 Commit
 Set Cn = Nothing
End Sub
#End If
