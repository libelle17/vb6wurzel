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
Private mvSV$ ' Sortiervariable = mv¸wnr & " " & mvName
Private mv‹WNr$
Private mv‹wNm$
Private mvName$
Private mvPat_ID&
Private mvobDMPInfo%
Public gew‰hlt%
Private mvarIndex As Long 'lokale Kopie
Public Property Let Index(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.index = 5
    mvarIndex = vData
End Property

Public Property Get Index() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.index
    Index = mvarIndex
End Property
Public Property Get SortVar()
 SortVar = mvSV
End Property

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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in istKleinerAls/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function '  istKleinerAls
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in unGleich/" + AnwPfad)
 Case vbAbort: Call MsgBox("Hˆre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' unGleich

Public Property Let ‹WNr(ByVal vData)
 If mv‹WNr <> vData Then
  mv‹WNr = vData
  mvSV = mv‹WNr & " " & mvName
 End If
End Property
Public Property Get ‹WNr() As Variant
 ‹WNr = mv‹WNr
End Property

Public Property Let ‹wNm(ByVal vData)
 If IsNull(vData) Then
  mv‹wNm = vNS
 Else
  mv‹wNm = vData
 End If
End Property
Public Property Get ‹wNm() As Variant
 ‹wNm = mv‹wNm
End Property

Public Property Let name(ByVal vData)
 If mvName <> vData Then
  mvName = vData
  mvSV = mv‹WNr & " " & mvName
 End If
End Property
Public Property Get name() As Variant
 name = mvName
End Property

Public Property Let Pat_id(ByVal vData)
 mvPat_ID = vData
End Property
Public Property Get Pat_id() As Variant
 Pat_id = mvPat_ID
End Property

Public Property Let obDMPInfo(ByVal vData)
 mvobDMPInfo = vData
End Property

Public Property Get obDMPInfo() As Variant
 obDMPInfo = mvobDMPInfo
End Property

