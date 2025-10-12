Attribute VB_Name = "Module1"
' benötigte API-Deklarationen
Private Declare Sub keybd_event Lib "User32" ( _
  ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
  ByVal dwExtraInfo As Long)
 
Private Declare Function VkKeyScan Lib "User32" _
  Alias "VkKeyScanA" ( _
  ByVal cChar As Byte) As Integer
 
Private Declare Function MapVirtualKey Lib "User32" _
  Alias "MapVirtualKeyA" ( _
  ByVal wCode As Long, _
  ByVal wMapType As Long) As Long
 
Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_EXTENDEDKEY = &H1
 
' Virtual KeyCodes
Public Enum eVirtualKeyCode
  VK_BAK = &H8
  VK_TAB = &H9
  VK_CLEAR = &HC
  VK_RETURN = &HD
  VK_SHIFT = &H10
  VK_CONTROL = &H11
  VK_MENU = &H12
  VK_PAUSE = &H13
  VK_CAPITAL = &H14
  VK_ESCAPE = &H1B
  VK_PRIOR = &H21
  VK_NEXT = &H22
  VK_END = &H23
  VK_HOME = &H24
  VK_LEFT = &H25
  VK_UP = &H26
  VK_RIGHT = &H27
  VK_DOWN = &H28
  VK_SELECT = &H29
  VK_SNAPSHOT = &H2C  ' NEU! Windows-Taste
  VK_INSERT = &H2D
  VK_DELETE = &H2E
  VK_HELP = &H2F
  VK_F1 = &H70
  VK_F2 = &H71
  VK_F3 = &H72
  VK_F4 = &H73
  VK_F5 = &H74
  VK_F6 = &H75
  VK_F7 = &H76
  VK_F8 = &H77
  VK_F9 = &H78
  VK_F10 = &H79
  VK_F11 = &H7A
  VK_F12 = &H7B
  VK_F13 = &H7C
  VK_F14 = &H7D
  VK_F15 = &H7E
  VK_F16 = &H7F
  VK_NUMLOCK = &H90
  VK_SCROLL = &H91
  VK_WIN = &H5B     ' NEU! Windows-Taste
  VK_APPS = &H5D    ' NEU! Taste für Kontextmenü
End Enum

' Text durch Simulieren von Tastenanschlägen
' an das aktive Control senden
Public Sub SendKeysEx(ByVal sText As String)
  Dim VK As eVirtualKeyCode
  Dim sChar As String
  Dim i As Integer
  Dim bShift As Boolean
  Dim bAlt As Boolean
  Dim bCtrl As Boolean
  Dim nScan As Long
  Dim nExtended As Long
 
  ' Jedes Zeichen einzeln senden
  For i = 1 To Len(sText)
    ' aktuelles Zeichen extrahieren
    sChar = Mid$(sText, i, 1)
 
    ' Sonderzeichen?
    bShift = False: bAlt = False: bCtrl = False
    If sChar = "{" Then
      If UCase$(Mid$(sText, i + 1, 9)) = "BACKSPACE" Then
        VK = VK_BAK
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "BS" Then
        VK = VK_BAK
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "BKSP" Then
        VK = VK_BAK
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "BREAK" Then
        VK = VK_PAUSE
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 8)) = "CAPSLOCK" Then
        VK = VK_CAPITAL
        i = i + 9
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "DELETE" Then
        VK = VK_DELETE
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "DEL" Then
        VK = VK_DELETE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "DOWN" Then
        VK = VK_DOWN
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "UP" Then
        VK = VK_UP
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "LEFT" Then
        VK = VK_LEFT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "RIGHT" Then
        VK = VK_RIGHT
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "END" Then
        VK = VK_END
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "ENTER" Then
        VK = VK_RETURN
        i = i + 6
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HOME" Then
        VK = VK_HOME
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "ESC" Then
        VK = VK_ESCAPE
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "HELP" Then
        VK = VK_HELP
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 6)) = "INSERT" Then
        VK = VK_INSERT
        i = i + 7
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "INS" Then
        VK = VK_INSERT
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 7)) = "NUMLOCK" Then
        VK = VK_NUMLOCK
        i = i + 8
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGUP" Then
        VK = VK_PRIOR
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "PGDN" Then
        VK = VK_NEXT
        i = i + 5
      ElseIf UCase$(Mid$(sText, i + 1, 10)) = "SCROLLLOCK" Then
        VK = VK_SCROLL
        i = i + 11
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "TAB" Then
        VK = VK_TAB
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F1" Then
        VK = VK_F1
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F2" Then
        VK = VK_F2
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F3" Then
        VK = VK_F3
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F4" Then
        VK = VK_F4
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F5" Then
        VK = VK_F5
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F6" Then
        VK = VK_F6
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F7" Then
        VK = VK_F7
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F8" Then
        VK = VK_F8
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 2)) = "F9" Then
        VK = VK_F9
        i = i + 3
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F10" Then
        VK = VK_F10
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F11" Then
        VK = VK_F11
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F12" Then
        VK = VK_F12
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F13" Then
        VK = VK_F13
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F14" Then
        VK = VK_F14
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F15" Then
        VK = VK_F15
        i = i + 4
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "F16" Then
        VK = VK_F16
        i = i + 4
 
      ' NEU! Windows-Taste
      ElseIf UCase$(Mid$(sText, i + 1, 3)) = "WIN" Then
        VK = VK_WIN
        i = i + 4
 
      ' NEU! Kontextmenü
      ElseIf UCase$(Mid$(sText, i + 1, 4)) = "APPS" Then
        VK = VK_APPS
        i = i + 5
 
      ' NEU! PrintScreen-Taste (DRUCK)
      ElseIf UCase$(Mid$(sText, i + 1, 5)) = "PRINT" Then
        VK = VK_SNAPSHOT
        i = i + 6
      End If
 
    ElseIf sChar = "+" Then
      ' Umschalttaste
      VK = VK_SHIFT
 
    ElseIf sChar = "%" Then
      ' ALT
      VK = VK_MENU
 
    ElseIf sChar = "^" Then
      ' STRG
      VK = VK_CONTROL
 
    Else
      ' Virtual KeyCode ermitteln...
      VK = VkKeyScan(Asc(sChar))
    End If
 
    nScan = MapVirtualKey(VK, 2)
    nExtended = 0
    If nScan = 0 Then nExtended = KEYEVENTF_EXTENDEDKEY
    nScan = MapVirtualKey(VK, 0)
 
    If VK <> VK_SHIFT Then
      ' Großbuchstabe...?
      bShift = (VK And &H100)
      bCtrl = (VK And &H200)
      bAlt = (VK And &H400)
      VK = (VK And &HFF)
    End If
 
    ' niederdrücken und wieder loslassen
    If bShift Then keybd_event VK_SHIFT, 0, 0, 0
    If bCtrl Then keybd_event VK_CONTROL, 0, 0, 0
    If bAlt Then keybd_event VK_MENU, 0, 0, 0
 
    keybd_event VK, nScan, nExtended, 0
    keybd_event VK, nScan, KEYEVENTF_KEYUP Or nExtended, 0
 
    ' Shift (Umsch)-Taste wieder loslassen
    If bShift Then keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
    If bCtrl Then keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    If bAlt Then keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
  Next i
End Sub

