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
' Klasse TMIniPropColl: Auflistung der Eigenschaften (TMIniProp) eines ini-Abschnitts
' ElemAdd: fügt eine vorhandene Eigenschaft an (optional mit Schlüssel)
' Aufruf in: TMIniPropColl.Kopiere, TMIniSectionColl.UEAdd
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ElemAdd/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'ElemAdd

' legt eine neue Eigenschaft Bezeichnung = Inhalt an und fügt sie an
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Add/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Add

' fügt der letzten Eigenschaft eine Untereigenschaft an
' (Hinweis: TMIniProp hat kein UnterEig und die Auflistung kein EAdd - wohl Überbleibsel, s. TMIniSectionColl.UEAdd)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UEAdd/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' UEAdd

' Element über Index oder (bei Text) über die Bezeichnung
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Item[PropertyGet]/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Item[PropertyGet]



' Zahl der Eigenschaften
' Aufruf in: TMIniPropColl.VergleicheMit
Public Property Get COUNT() As Long
 On Error GoTo fehler
 COUNT = mCol.COUNT
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Count/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Count


' entfernt eine Eigenschaft über Index oder Schlüssel
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Sub Remove(vntIndexKey As Variant)
 On Error GoTo fehler
 mCol.Remove vntIndexKey
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Remove/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Remove


' Aufzähler für For Each
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
 On Error GoTo fehler
    'Diese Eigenschaft ermöglicht, die Elemente der Auflistung mit For...Each zu durchlaufen.
 Set NewEnum = mCol.[_NewEnum]
 Exit Property
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in NewEnum/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' NewEnum


' legt beim Erzeugen die Auflistung an
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
 Set mCol = New Collection
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Initialize/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Initialize


' gibt die Auflistung beim Beenden frei
' Aufruf in: Ereignisprozedur
Private Sub Class_Terminate()
 Set mCol = Nothing
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Class_terminate/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_terminate

' liefert eine Kopie der Auflistung (mit kopierten Elementen)
' Aufruf in: TMIniSection.Kopiere
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Kopiere/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Kopiere

' wahr (-1), wenn Vg gleich viele Eigenschaften mit gleichen Bezeichnungen und Inhalten in gleicher Reihenfolge hat
' Aufruf in: TMIniSection.VergleicheMit
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in VergleicheMit/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VergleicheMit
