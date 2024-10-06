VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMIniPropColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2
Attribute VB_Description = "Sammlung von Eigenschaften"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TMEigS"
Attribute VB_Ext_KEY = "Member0" ,"TMEigS"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Lokale Variable für eine Auflistung.
Private mCol As Collection
Public Function ElemAdd(Eig As TMIniProp, Optional skey$) As TMIniProp
  On Error GoTo fehler
    If Len(skey) = 0 Then
        mCol.Add Eig
    Else
        mCol.Add Eig, skey
    End If
    'das erstellte Objekt zurückgeben
    Set ElemAdd = Eig
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ElemAdd/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'ElemAdd
Public Function Add(Bezeichnung$, Inhalt$, Optional skey$) As TMIniProp
    'neues Objekt erstellen
    On Error GoTo fehler
    Dim objNewMember As New TMIniProp
    'die Eigenschaften festlegen, die der Methode übergeben werden
    objNewMember.Bezeichnung = Bezeichnung
    objNewMember.Inhalt = Inhalt
    If Len(skey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, skey
    End If
    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Add/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Add
Public Function UEAdd(Bezeichnung$, Inhalt$)
 On Error GoTo fehler
 Dim objNewMember As New TMIniProp
    objNewMember.Bezeichnung = Bezeichnung
    objNewMember.Inhalt = Inhalt
    If mCol.Item(mCol.COUNT).UnterEig Is Nothing Then Set mCol.Item(mCol.COUNT).UnterEig = New TMIniPropColl
    mCol.Item(mCol.COUNT).UnterEig.EAdd objNewMember
    Set UEAdd = objNewMember
    Set objNewMember = Nothing
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UEAdd/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' UEAdd

Public Property Get Item(vntIndexKey As Variant) As TMIniProp
Attribute Item.VB_UserMemId = 0
    'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
    'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
    'Daher wird der Parameter als Variant deklariert.
    'Syntax: SET foo = x.Item(xyz) OR SET foo = x.Item(5)
  Dim i&
  On Error GoTo fehler
  If VarType(vntIndexKey) = 8 Then
   For i = 1 To mCol.COUNT
    If mCol.Item(i).Bezeichnung = vntIndexKey Then
     Set Item = mCol.Item(i)
     Exit Property
    End If
   Next i
  Else
   Set Item = mCol(vntIndexKey)
  End If
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Item[PropertyGet]/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Item[PropertyGet]



Public Property Get COUNT() As Long
 On Error GoTo fehler
 COUNT = mCol.COUNT
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Count/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Count


Public Sub Remove(vntIndexKey As Variant)
 On Error GoTo fehler
 mCol.Remove vntIndexKey
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Remove/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Remove


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
 On Error GoTo fehler
    'Diese Eigenschaft ermöglicht, die Elemente der Auflistung mit For...Each zu durchlaufen.
 Set NewEnum = mCol.[_NewEnum]
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in NewEnum/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' NewEnum


Private Sub Class_Initialize()
 Set mCol = New Collection
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Initialize/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Initialize


Private Sub Class_Terminate()
 Set mCol = Nothing
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_terminate/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_terminate

Public Function Kopiere() As TMIniPropColl
 Dim i&
 On Error GoTo fehler
 Dim tmpPC As New TMIniPropColl
 For i = 1 To mCol.COUNT
  Call tmpPC.ElemAdd(mCol(i).Kopiere)
 Next i
 Set Kopiere = tmpPC
 Set tmpPC = Nothing
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Kopiere/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Kopiere

Public Function VergleicheMit(Vg As TMIniPropColl)
 Dim i&, imax&
 Dim iMem As TMIniProp
 On Error GoTo fehler
 imax = mCol.COUNT
 If Vg Is Nothing Then Exit Function
 If mCol.COUNT = Vg.COUNT Then
  For i = 1 To imax
   Set iMem = mCol(i)
   If Not iMem.VergleicheMit(Vg(i)) Then Exit Function
  Next i
  VergleicheMit = -1
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in VergleicheMit/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VergleicheMit
