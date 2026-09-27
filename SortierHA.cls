VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortierHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text ' 1.10.25
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Private mvarInhalt AS Variant 'lokale Kopie
Private mvÜWNr$
Private mvÜwNm$
Private mvFax$
Private mvobDMPInfo%
Public Zahl&
Public gewählt%
Private mvarIndex As Long 'lokale Kopie

' Klasse SortierHA: Sortierelement für Hausärzte (Überweisernummer ÜWNr, Name ÜwNm, Fax, obDMPInfo), sortiert nach Name
' Index: Position des Elements (setzen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Let Index(ByVal vData&)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.index = 5
    mvarIndex = vData
End Property

' Index: Position des Elements (lesen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get Index() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.index
    Index = mvarIndex
End Property

' Sortierwert (Name des Hausarztes)
' Aufruf in: SortierEintr.istKleinerAls, SortierEintr.unGleich, SortierHA.istKleinerAls, SortierHA.unGleich, SortierPatZuHA.istKleinerAls,
'   SortierPatZuHA.unGleich
Public Property Get SortVar()
 SortVar = mvÜwNm
End Property

' Sortiervergleich nach Name: wahr, wenn dieses Element vor E2 (Element oder Wert) einzuordnen ist
' Aufruf in: SortierListe.Eingrenz
Public Function istKleinerAls(ByRef E2) As Boolean
 On Error GoTo fehler
 If IsObject(E2) Then
  If mvÜwNm < E2.SortVar Then istKleinerAls = -1
 Else
  If mvÜwNm < E2 Then istKleinerAls = -1
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in istKleinerAls/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function '  istKleinerAls

' wahr, wenn sich der Name von E2 (Element oder Wert) unterscheidet
' Aufruf in: SortierListe.GetItem, SortierListe.sCAdd, SortierListe.SuchItem
Public Function unGleich(ByRef E2) As Boolean
 On Error GoTo fehler
 If IsObject(E2) Then
  If mvÜwNm <> E2.SortVar Then unGleich = True
 Else
  If mvÜwNm <> E2 Then unGleich = True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + CStr(nz(Err.Source, "")) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in unGleich/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' unGleich

' Überweiser-/Arztnummer (setzen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click
Public Property Let ÜWNr(ByVal vData)
 If mvÜWNr <> vData Then
  mvÜWNr = vData
 End If
End Property

' Überweiser-/Arztnummer (lesen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click
Public Property Get ÜWNr() As Variant
 ÜWNr = mvÜWNr
End Property

' Name des Hausarztes (setzen, Null wird zu Leerstring)
' Aufruf in: PatListe.DMPFüll, PatListe.doStart
Public Property Let ÜwNm(ByVal vData)
 If IsNull(vData) Then
  mvÜwNm = vNS
 Else
  mvÜwNm = vData
 End If
End Property

' Name des Hausarztes (lesen)
' Aufruf in: PatListe.DMPFüll, PatListe.doStart
Public Property Get ÜwNm() As Variant
 ÜwNm = mvÜwNm
End Property

' Faxnummer (setzen)
' Aufruf in: PatListe.DMPFüll, PatListe.doStart
Public Property Let fax(ByVal vData)
 If mvFax <> vData Then
  mvFax = vData
 End If
End Property

' Faxnummer (lesen)
' Aufruf in: PatListe.DMPFüll, PatListe.doStart
Public Property Get fax() As Variant
 fax = mvFax
End Property

' ob der Arzt DMP-Informationen erhält (setzen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll
Public Property Let obDMPInfo(ByVal vData)
 If mvobDMPInfo <> vData Then
  mvobDMPInfo = vData
 End If
End Property

' ob der Arzt DMP-Informationen erhält (lesen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll
Public Property Get obDMPInfo() As Variant
 obDMPInfo = mvobDMPInfo
End Property
