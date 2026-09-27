VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cValuePairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The private collection used to hold the real data
Private m_cValuePairs As Collection

' Klasse cValuePairs: Auflistung von Name/Wert-Paaren (cValuePair), z.B. der Umgebungsvariablen (SystemEnvironment.List)
' Class_Initialize: legt beim Erzeugen die Auflistung an
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_cValuePairs = New Collection
End Sub

' Add a new cValuePair item to the collection

' fügt ein Paar an (optional mit Schlüssel)
' Aufruf in: System.EnumerateVariables
Friend Sub Add(newItem As cValuePair, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    ' TO DO: initialize new item's properties here
    ' ...
    ' add to the private collection
    m_cValuePairs.Add newItem, Key
End Sub

' Remove an item FROM the collection

' entfernt ein Paar über Index oder Schlüssel
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Friend Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_cValuePairs.Remove Index
End Sub

' Return a cValuePair item FROM the collection

' Paar über Index oder Schlüssel
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Function Item(Index As Variant) As cValuePair
Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_cValuePairs.Item(Index)
End Function

' Return the number of items in the collection

' Zahl der Paare
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Property Get COUNT() As Long
Attribute COUNT.VB_Description = "Returns the number of members in a collection"
    COUNT = m_cValuePairs.COUNT
End Property

' Remove all items FROM the collection

' leert die Auflistung
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Friend Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_cValuePairs = New Collection
End Sub

' Implement support for enumeration (For Each)

' Aufzähler für For Each
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_cValuePairs.[_NewEnum]
End Function

