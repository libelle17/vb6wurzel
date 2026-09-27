VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrEnvironment As String
Private Const REGISTRY_PATH = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
Private oRegistry As Registry

Private Declare Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, _
    ByVal lParam As String, ByVal fuFlags As Long, _
    ByVal uTimeout As Long, lpdwResult As Long) As Long
    
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const WM_SETTINGCHANGE As Long = &H1A

Private mValuePairs As cValuePairs
' Klasse System (SystemEnvironment): liest und setzt die systemweiten Umgebungsvariablen in der Registry
' (HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment)
' EnumerateVariables: liest alle Variablen als Name/Wert-Paare nach mValuePairs
' Aufruf in: System.List
Private Sub EnumerateVariables()
    Dim sKeys() As String
    Dim iKeyCount As Long
    Dim x As Long
    Dim oValuePair As cValuePair
    
    Set mValuePairs = New cValuePairs
    
    oRegistry.EnumerateVALUES sKeys(), iKeyCount
    For x = 1 To iKeyCount
        oRegistry.ValueKey = sKeys(x)
        
        Set oValuePair = New cValuePair
        oValuePair.Variable = sKeys(x)
        oValuePair.Value = oRegistry.Value
        
        mValuePairs.Add oValuePair, oValuePair.Variable
        Set oValuePair = Nothing
    Next
    
End Sub

' Wert der Umgebungsvariablen strName (lesen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get Environment(ByVal strName As String) As String
    oRegistry.ValueKey = strName
    Environment = oRegistry.Value
End Property

' Wert der Umgebungsvariablen strName (als REG_SZ setzen; Benachrichtigung der Programme ist abgeschaltet)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Let Environment(ByVal strName As String, ByVal strValue As String)
    oRegistry.ValueKey = strName
    oRegistry.ValueType = REG_SZ
    oRegistry.Value = strValue
    
'    BroadcastEnvironmentChange
End Property

' benachrichtigt alle Fenster über geänderte Umgebungsvariablen (WM_SETTINGCHANGE)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Private Sub BroadcastEnvironmentChange()
    Dim lngReturnValue As Long

    SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, 0&, _
        "Environment", SMTO_ABORTIFHUNG, 5000&, lngReturnValue
End Sub


' alle Umgebungsvariablen als Name/Wert-Paare (beim ersten Zugriff gelesen)
' Aufruf in: (kein Aufruf in DateiLese.vbp gefunden)
Public Property Get List() As cValuePairs
    If mValuePairs Is Nothing Then
        EnumerateVariables
    End If
    
    Set List = mValuePairs
End Property

' öffnet beim Erzeugen den Registry-Schlüssel der Umgebungsvariablen
' Aufruf in: Ereignisprozedur
Private Sub Class_Initialize()
    Set oRegistry = New Registry
    oRegistry.ClassKey = HKEY_LOCAL_MACHINE
    oRegistry.SectionKey = REGISTRY_PATH
End Sub


' gibt Registry-Objekt und Liste beim Beenden frei
' Aufruf in: Ereignisprozedur
Private Sub Class_Terminate()
    Set oRegistry = Nothing
    Set mValuePairs = Nothing
End Sub


