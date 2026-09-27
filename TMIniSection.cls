VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMIniSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarBezeichnung$ 'lokale Kopie
Private mvarUnterEig As TMIniPropColl 'lokale Kopie

' Klasse TMIniSection: ein Abschnitt einer Turbomed-ini-Datei (Bezeichnung und Eigenschaften UnterEig)
' UnterEig: Eigenschaften des Abschnitts (setzen)
' Aufruf in: TMIniSection.Item, TMIniSection.Kopiere, TMIniSection.VergleicheMit
Public Property Set UnterEig(ByVal vData As TMIniPropColl)
 On Error GoTo fehler
 Set mvarUnterEig = vData
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UnterEig[PropertySet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' UnterEig[PropertySet]

' Eigenschaften des Abschnitts (lesen)
' Aufruf in: TMIniSection.Item, TMIniSection.Kopiere, TMIniSection.VergleicheMit
Public Property Get UnterEig() As TMIniPropColl
 On Error GoTo fehler
 Set UnterEig = mvarUnterEig
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UnterEig[PropertyGet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' UnterEig[PropertyGet]

' Name des Abschnitts (setzen)
' Aufruf in: TMIniSection.Kopiere, TMIniSection.VergleicheMit, TMIniSectionColl.Add, TMIniSectionColl.Item
Public Property Let Bezeichnung(ByVal vData$)
 On Error GoTo fehler
 mvarBezeichnung = vData
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Bezeichnung[PropertyLet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Bezeichnung[PropertyLet]


' Name des Abschnitts (lesen)
' Aufruf in: TMIniSection.Kopiere, TMIniSection.VergleicheMit, TMIniSectionColl.Add, TMIniSectionColl.Item
Public Property Get Bezeichnung$()
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Bezeichnung
 On Error GoTo fehler
 Bezeichnung = mvarBezeichnung
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Bezeichnung[PropertyGet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Bezeichnung[PropertyGet]

' Eigenschaft des Abschnitts über Index oder Bezeichnung
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get Item(vntIndexKey As Variant) As TMIniProp
 On Error GoTo fehler
 Set Item = UnterEig.Item(vntIndexKey)
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Item[PropertyGet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Item[PropertyGet]

' liefert eine Kopie des Abschnitts samt Eigenschaften
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Function Kopiere() As TMIniSection
 On Error GoTo fehler
 Dim tmpS As New TMIniSection
 tmpS.Bezeichnung = mvarBezeichnung
 Set tmpS.UnterEig = mvarUnterEig.Kopiere
 Set Kopiere = tmpS
 Set tmpS = Nothing
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Kopiere/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Kopiere

' wahr (-1), wenn Vg gleich heißt und gleiche Eigenschaften hat
' Aufruf in: TMIniSectionColl.VergleicheMit
Public Function VergleicheMit(Vg As TMIniSection)
 On Error GoTo fehler
 If Vg.Bezeichnung = mvarBezeichnung Then
  If mvarUnterEig.VergleicheMit(Vg.UnterEig) Then
   VergleicheMit = -1
  End If
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in VergleicheMit/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VergleicheMit
