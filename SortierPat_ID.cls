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
Public Property Let Index(ByVal vData As Long)
    mvIndex = vData
End Property
Public Property Get Index() As Long
    Index = mvIndex
End Property

Public Function istKleinerAls(ByRef E2) As Boolean
Attribute istKleinerAls.VB_Description = "Wahr, wenn das erste Element größer ist"
 If mvPat_ID < E2.Pat_id Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
End Function
Public Function unGleich(ByRef E2) As Boolean
 If mvPat_ID <> E2.Pat_id Then unGleich = True
End Function
Public Property Let Pat_id(ByVal vData&)
    mvPat_ID = vData
End Property
Public Property Get Pat_id&()
    Pat_id = mvPat_ID
End Property

Public Property Let Knz(ByVal vData&)
    mvKnz = vData
End Property
Public Property Get Knz&()
    Knz = mvKnz
End Property

Public Property Let Zeitpunkt(ByVal vData#)
    mvZeitpunkt = vData
End Property
Public Property Get Zeitpunkt#()
    Zeitpunkt = mvZeitpunkt
End Property


