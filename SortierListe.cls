VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortierListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"SortierElem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Lokale Variable für eine Auflistung.
Private mCol As Collection
'Const sIZahl& = 2 ' = Zahl der verschiedenen sortierten Kollektionen, muss im jeweiligen Projekt definiert werden
'Public sC(sIZahl - 1) AS Collection
Public sCa&, sCe& ' 0 = Archiv, 1 = Eingelesene

' Klasse SortierListe: sortierte Auflistung (mCol) beliebiger Sortierelemente, die istKleinerAls/unGleich bereitstellen
' (Binärsuche über den Bereich sCa..sCe)
' GetItem: liefert das Element, das SI entspricht (unGleich = False), sonst Nothing
' Aufruf in: Importiere.alleSpeichern, Importiere.laborparameterSpeichern, vonMo.doPatvonMO, vonMo.MODmpreihe, ZielDBFunktionen.hollabor
Public Function GetItem(SI) As Object '$, ind%)
  On Error GoTo fehler
'  Call sCInit(ind)
  Call scinit
  ' See WHERE the item belongs.
  Call Eingrenz(SI)
  Set GetItem = Nothing
  If sCe < sCa Then Else If mCol(sCa).unGleich(SI) Then Else Set GetItem = mCol(sCa)
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SuchItem/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SuchItem

' wahr (-1), wenn ein SI entsprechendes Element in der Liste steht
' Aufruf in: PatListe.AlleMark, PatListe.Form_Load, PatListe.MFG_Click
Public Function SuchItem%(SI) '$, ind%)
  On Error GoTo fehler
'  Call sCInit(ind)
  Call scinit
  ' See WHERE the item belongs.
  Call Eingrenz(SI)
  If sCe < sCa Then Else If mCol(sCa).unGleich(SI) Then Else SuchItem = -1
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SuchItem/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SuchItem

' Binärsuche: engt den Bereich sCa..sCe rekursiv ein, bis sCa die Einfügeposition für new_item ist
' (erstes Element, das nicht kleiner als new_item ist)
' Aufruf in: SortierListe.GetItem, SortierListe.sCAdd, SortierListe.SuchItem
Public Function Eingrenz(new_item) ', ByRef A&, ByRef E&)
 Dim Mitte&, diff&, i&
 On Error GoTo fehler
 diff = sCe - sCa
 If diff < 6 Then
  For i = IIf(sCa = 0, 1, sCa) To sCe
    If Not mCol(i).istKleinerAls(new_item) Then Exit For
  Next i
  sCa = i
  Exit Function
 End If
 Mitte = sCa + 0.5 * diff
 If mCol(Mitte).istKleinerAls(new_item) Then
  sCa = Mitte
 Else
  sCe = Mitte
 End If
 Call Eingrenz(new_item)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Eingrenz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Eingrenz

' setzt den Suchbereich auf die ganze Liste (sCa = 0, sCe = Anzahl) und legt mCol bei Bedarf an
' Aufruf in: SortierListe.Class_Initialize, SortierListe.GetItem, SortierListe.sCAdd, SortierListe.SuchItem
Public Function scinit()
 On Error GoTo fehler
 sCa = 0
 If mCol Is Nothing Then
  Set mCol = New Collection
 Else
  sCe = mCol.COUNT
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in sCInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' sCInit

' fügt new_item an der sortierten Position ein; bei obDistinct wird ein gleiches vorhandenes Element nicht doppelt
' eingefügt, sondern dessen Zähler mvAnzahl (und bei anderer Pat_id mvPatZahl) erhöht
' Aufruf in: Haupt.VergleichTab, Haupt.ÜbertragFormulare, Importiere.alleSpeichern, Importiere.doMedklassT, Importiere.doTabVorb,
'   Importiere.laborparameterSpeichern, Importiere.laborparvorladen, Importiere.medartenhier, ImportiereLabor.LaborDirektImport, modMemoScan.MemoScanToList,
'   PatListe.DMPFüll, PatListe.Form_Load, vonMo.MOKatLaden, vonMo.zeigmosystem, ZielDBFunktionen.dododoPorto, ZielDBFunktionen.hollabor
Public Sub sCAdd(ByVal new_item, Optional obDistinct)
  Dim i&, AddRes
  On Error GoTo fehler
  If IsMissing(obDistinct) Then obDistinct = 0
  Call scinit
  ' See WHERE the item belongs.
  Call Eingrenz(new_item)
  ' Insert the item.
  If sCa > mCol.COUNT Then
   ' Add at the end.
   Set AddRes = Add(new_item, 0)
  Else
   ' Add at the right position.
   If Not obDistinct Or mCol(sCa).unGleich(new_item) Then
    Set AddRes = Add(new_item, sCa)
   Else
    Set AddRes = mCol(sCa)
   End If
  End If
  If Not AddRes Is Nothing Then
   On Error Resume Next
   AddRes.mvAnzahl = AddRes.mvAnzahl + 1
   If new_item.Pat_id <> AddRes.Pat_id Or AddRes.mvPatZahl = 0 Then
    AddRes.mvPatZahl = AddRes.mvPatZahl + 1
   End If
  End If
  Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in SCAdd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub 'AddIdem

' hängt Element an (skey = 0) bzw. fügt es vor Position skey ein und gibt es zurück
' Aufruf in: ImportiereLabor.LaborErgPatId, SortierListe.sCAdd
Public Function Add(Element, Optional skey&)
  On Error GoTo fehler
    If skey = 0 Then
        mCol.Add Element
    Else
        mCol.Add Element, , skey
    End If
    If IsObject(Element) Then
      Set Add = Element
    Else
      Add = Element
    End If
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Add/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Add

' Element über Index oder Schlüssel
' Aufruf in: Haupt.VergleichTab, Haupt.ÜbertragFormulare, Importiere.medklass2, ImportiereLabor.LaborDirektImport, ImportiereLabor.LaborErgPatId,
'   PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart, PatListe.MFG_Click, PatListe.zähleDMPPat, ZielDBFunktionen.dododoPorto
Public Property Get Item(vntIndexKey As Variant)
    'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
    'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
    'Daher wird der Parameter als Variant deklariert.
    'Syntax: SET foo = x.Item(xyz) OR SET foo = x.Item(5)
  On Error GoTo fehler
  If IsObject(mCol(vntIndexKey)) Then
    Set Item = mCol(vntIndexKey)
  Else
    Item = mCol(vntIndexKey)
  End If
'  On Error Resume Next
'  Set Item = mCol(vntIndexKey)
'  If Err.Number <> 0 Then
'   Item = mCol(vntIndexKey)
'  End If
 Exit Property
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Item/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Get Item

' Zahl der Elemente
' Aufruf in: Haupt.VergleichTab, Haupt.ÜbertragFormulare, Importiere.doTabVorb, Importiere.laborparameterSpeichern, Importiere.laborparvorladen,
'   Importiere.medklass2, ImportiereLabor.LaborDirektImport, ImportiereLabor.LaborErgPatId, PatListe.AlleMark, PatListe.DMPFüll, PatListe.doStart,
'   PatListe.MFG_Click, PatListe.zähleDMPPat, vonMo.MOKatLaden, ZielDBFunktionen.dododoPorto
Public Property Get COUNT() As Long
    'Wird verwendet, wenn die Anzahl der Elemente
    'in der Auflistung ermittelt werden soll.
    'Syntax: Debug.Print x.Count
    On Error GoTo fehler
    COUNT = mCol.COUNT
 Exit Property
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Get_COUNT()/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Get Count

' entfernt ein Element über Index oder Schlüssel
' Aufruf in: ImportiereLabor.LaborErgPatId
Public Sub Remove(vntIndexKey As Variant)
    'Wird verwendet, wenn ein Element aus der Auflistung
    'entfernt werden soll. vntIndexKey enthält entweder
    'den Index oder den Schlüssel. Daher wird der Parameter
    'als Variant deklariert.
    'Syntax: x.Remove(xyz)

    On Error GoTo fehler
    mCol.Remove vntIndexKey
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Remove/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

' Aufzähler für For Each
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get NewEnum() As IUnknown
    'Diese Eigenschaft ermöglicht, die Elemente der
    'Auflistung mit For...Each zu durchlaufen.
    On Error GoTo fehler
    Set NewEnum = mCol.[_NewEnum]
 Exit Property
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in NewEnum/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Property ' Get NewEnum

' legt beim Erzeugen die Auflistung an
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
    'Erstellt die Auflistung, wenn die klasse erstellt wird.
    'Set mCol = New Collection
    On Error GoTo fehler
    Call scinit
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Initialize/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Initialize

' gibt die Auflistung beim Beenden frei
' Aufruf in: Ereignisprozedur
Private Sub Class_Terminate()
    'Entfernt die Auflistung aus dem Speicher, wenn die klasse beendet wird.
    On Error GoTo fehler
    Set mCol = Nothing
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Class_Terminate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Class_Terminate

