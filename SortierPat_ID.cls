VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortierPat_ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Zu sortierende Einheit"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvPat_ID& 'lokale Kopie
Private mvZeitpunkt#
Private mvKnz%
Private mvIndex& 'lokale Kopie
' Klasse SortierPat_ID: Sortierelement nach Pat_id (mit Kennzeichen Knz und Zeitpunkt), verwendet mit SortierListe
' Index: Position des Elements (setzen)
' 27.9.26: Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Let Index(ByVal vData As Long)
    mvIndex = vData
End Property

' Index: Position des Elements (lesen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get Index() As Long
    Index = mvIndex
End Property

' Sortiervergleich nach Pat_id: wahr, wenn dieses Element vor E2 einzuordnen ist
' Aufruf in: SortierListe.Eingrenz
Public Function istKleinerAls(ByRef E2) As Boolean
Attribute istKleinerAls.VB_Description = "Wahr, wenn das erste Element größer ist"
 If mvPat_ID < E2.Pat_id Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
End Function

' wahr, wenn sich die Pat_id von der in E2 unterscheidet
' Aufruf in: SortierListe.GetItem, SortierListe.sCAdd, SortierListe.SuchItem
Public Function unGleich(ByRef E2) As Boolean
 If mvPat_ID <> E2.Pat_id Then unGleich = True
End Function

' Patientennummer (setzen)
' Aufruf in: ImportiereLabor.LaborErgPatId, SortierPat_ID.istKleinerAls, SortierPat_ID.unGleich
Public Property Let Pat_id(ByVal vData&)
    mvPat_ID = vData
End Property

' Patientennummer (lesen)
' Aufruf in: ImportiereLabor.LaborErgPatId, SortierPat_ID.istKleinerAls, SortierPat_ID.unGleich
Public Property Get Pat_id&()
    Pat_id = mvPat_ID
End Property

' Kennzeichen (setzen)
' Aufruf in: ImportiereLabor.LaborErgPatId
Public Property Let Knz(ByVal vData&)
    mvKnz = vData
End Property

' Kennzeichen (lesen)
' Aufruf in: ImportiereLabor.LaborErgPatId
Public Property Get Knz&()
    Knz = mvKnz
End Property

' Zeitpunkt (setzen)
' Aufruf in: ImportiereLabor.LaborErgPatId
Public Property Let Zeitpunkt(ByVal vData#)
    mvZeitpunkt = vData
End Property

' Zeitpunkt (lesen)
' Aufruf in: ImportiereLabor.LaborErgPatId
Public Property Get Zeitpunkt#()
    Zeitpunkt = mvZeitpunkt
End Property


