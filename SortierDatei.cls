VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortierDatei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Zu sortierende Einheit"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvFile As File
Private mvDatID&
Private mvIndex& 'lokale Kopie
Public Property Let Index(ByVal vData As Long)
    mvIndex = vData
End Property
Public Property Get Index() As Long
    Index = mvIndex
End Property

Public Function istKleinerAls(ByRef E2) As Boolean
Attribute istKleinerAls.VB_Description = "Wahr, wenn das erste Element größer ist"
 If mvFile.DateLastModified < E2.File.DateLastModified Then
  istKleinerAls = -1
 Else
  istKleinerAls = 0
 End If
End Function
Public Function unGleich(ByRef E2) As Boolean
 If mvFile.DateLastModified <> E2.File.DateLastModified Then unGleich = True
End Function
Public Property Set File(ByVal vData)
    Set mvFile = vData
End Property
Public Property Get File()
    Set File = mvFile
End Property

Public Property Let File(mvFile)
    Set File = mvFile
End Property

Public Property Let DatID(ByVal vData&)
    mvDatID = vData
End Property
Public Property Get DatID&()
    DatID = mvDatID
End Property

