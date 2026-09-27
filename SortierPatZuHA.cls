VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortierPatZuHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
'Private mvarInhalt AS Variant 'lokale Kopie
Private mvSV$ ' Sortiervariable = mvüwnr & " " & mvName
Private mvÜWNr$
Private mvÜwNm$
Private mvName$
Private mvPat_ID&
Private mvobDMPInfo%
Public sCa& ' Index des zugehordneten Arztes (HASL)
Public gewählt%
Private mvarIndex As Long 'lokale Kopie
' Klasse SortierPatZuHA: Sortierelement Patient zu Hausarzt, sortiert nach "ÜWNr Name" (mvSV)
' Index: Position des Elements (setzen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Let Index(ByVal vData As Long)
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

' Sortierwert "Überweisernummer Name"
' Aufruf in: SortierEintr.istKleinerAls, SortierEintr.unGleich, SortierHA.istKleinerAls, SortierHA.unGleich, SortierPatZuHA.istKleinerAls,
'   SortierPatZuHA.unGleich
Public Property Get SortVar()
 SortVar = mvSV
End Property

' Sortiervergleich nach SortVar: wahr, wenn dieses Element vor E2 (Element oder Wert) einzuordnen ist
' Aufruf in: SortierListe.Eingrenz
Public Function istKleinerAls(ByRef E2) As Boolean
 On Error GoTo fehler
 If IsObject(E2) Then
  If mvSV < (E2.SortVar) Then istKleinerAls = True
 Else
  If mvSV < E2 Then istKleinerAls = True
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

' wahr, wenn sich SortVar von E2 (Element oder Wert) unterscheidet
' Aufruf in: SortierListe.GetItem, SortierListe.sCAdd, SortierListe.SuchItem
Public Function unGleich(ByRef E2) As Boolean
 On Error GoTo fehler
 If IsObject(E2) Then
  If mvSV <> E2.SortVar Then unGleich = True
 Else
  If mvSV <> E2 Then unGleich = True
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

' Überweiser-/Arztnummer (setzen, aktualisiert den Sortierwert)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click
Public Property Let ÜWNr(ByVal vData)
 If mvÜWNr <> vData Then
  mvÜWNr = vData
  mvSV = mvÜWNr & " " & mvName
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

' Patientenname (setzen, aktualisiert den Sortierwert)
' Aufruf in: PatListe.DMPFüll
Public Property Let name(ByVal vData)
 If mvName <> vData Then
  mvName = vData
  mvSV = mvÜWNr & " " & mvName
 End If
End Property

' Patientenname (lesen)
' Aufruf in: PatListe.DMPFüll
Public Property Get name() As Variant
 name = mvName
End Property

' Patientennummer (setzen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click
Public Property Let Pat_id(ByVal vData)
 mvPat_ID = vData
End Property

' Patientennummer (lesen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click
Public Property Get Pat_id() As Variant
 Pat_id = mvPat_ID
End Property

' ob der Arzt DMP-Informationen erhält (setzen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll
Public Property Let obDMPInfo(ByVal vData)
 mvobDMPInfo = vData
End Property

' ob der Arzt DMP-Informationen erhält (lesen)
' Aufruf in: PatListe.AlleMark, PatListe.DMPFüll
Public Property Get obDMPInfo() As Variant
 obDMPInfo = mvobDMPInfo
End Property

