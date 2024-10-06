VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShellLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------
' Klassenmodul cShellLink
'
' Dient dem Erzeugen und dem Auslesen von Windows-
' Verknüpfungsdateien (*.lnk)
'
' Voraussetzung: Verweis auf die Typbibliothek SHELLLNK.TLB!
' Sie finden diese Datei auf Ihrer Visual Basic- oder
' Visual Studio-CD:
' Visual Basic-CD: \Tools\Unsupprt\ShellLnk
' Visual Studio-CD 3: \common\Tools\vb\Unsupprt\ShellLnk
' Die Typbibliothek wird nur für das Compilieren benötigt:
' Eine Auslieferung mit Ihrer Anwendung ist unnötig!
' -----------------------------------------------------------
' Um Informationen einer bestehenden Verknüpfungsdatei zu
' ermitteln, belegen Sie die FileName-Eigenschaft dieser
' Klasse entsprechend und führen Sie dann die Funktion
' LoadLink aus. Um eine Verknüpfungsdatei zu erstellen,
' belegen Sie die gewünschten Eigenschaften mit Werten,
' vergeben Sie an die FileName-Eigenschaft einen Pfad für
' die neue Verknüpfung und führen Sie die Funktion SaveLink
' aus.
' -----------------------------------------------------------
' Einen einfacheren Weg, Verküpfungen weniger flexibel zu
' erstellen, finden Sie in der Microsoft Knowledge Base:
' HOWTO: Create Shortcuts (Shell Links) within Windows
' <A href="http://support.microsoft.com/default.aspx?scid=kb;DE;155303">http://support.microsoft.com/default.aspx?scid=kb;DE;155303</A>
' -----------------------------------------------------------
' Copyright (c) 2003 by Mathias Schiffer, AixSoft Software.
' Downloadmöglichkeit: <A href="http://www.aixsoft.de/msdn/quickie">http://www.aixsoft.de/msdn/quickie</A>
' -----------------------------------------------------------
' ---- BENÖTIGTE DEKLARATIONEN ----
' Werte für die Eigenschaft StartupWindowState
Public Enum StartupWindowStates
  swsNormal = 1
  swsMaximized = 7
  swsMinimized = 3
End Enum
' Eigenschaftenvariable
Private mDescription  As String
Private mPath   As String
Private mRelativePath As String
Private mWorkingDirectory   As String
Private mArguments As String
Private mStartupWindowState As StartupWindowStates
Private mHotkey As VBRUN.KeyCodeConstants
Private mHotkeyModifiers As VBRUN.ShiftConstants
Private mIconPath  As String
Private mIconNumber   As Long
Private mFileName  As String
' API-Deklarationen zur Ermittlung spezieller Pfade
Public Enum ShellFolders
  CSIDL_DESKTOP = &H0&
  CSIDL_PROGRAMS = &H2&
  CSIDL_CONTROLS = &H3&
  CSIDL_PRINTERS = &H4&
  CSIDL_PERSONAL = &H5&
  CSIDL_FAVORITES = &H6&
  CSIDL_STARTUP = &H7&
  CSIDL_RECENT = &H8&
  CSIDL_SENDTO = &H9&
  CSIDL_BITBUCKET = &HA&
  CSIDL_STARTMENU = &HB&
  CSIDL_DESKTOPDIRECTORY = &H10&
  CSIDL_DRIVES = &H11&
  CSIDL_NETWORK = &H12&
  CSIDL_NETHOOD = &H13&
  CSIDL_FONTS = &H14&
  CSIDL_TEMPLATES = &H15&
  CSIDL_COMMON_STARTMENU = &H16&
  CSIDL_COMMON_PROGRAMS = &H17&
  CSIDL_COMMON_STARTUP = &H18&
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19&
  CSIDL_APPDATA = &H1A&
  CSIDL_PRINTHOOD = &H1B&
End Enum
Private Type SHITEMID
  cb As Long
  abID As Byte
End Type
Private Type ITEMIDLIST
  mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation _
  Lib "shell32.dll" ( _
  ByVal hwndOwner As Long, _
  ByVal nFolder As Long, _
  ByRef pidl As ITEMIDLIST _
  ) As Long
Private Declare Function SHGetPathFromIDList _
  Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
  ByVal pidl As Long, _
  ByVal pszPath As String _
  ) As Long
' ---- EIGENSCHAFTENPROZEDUREN ----
Public Property Let FileName(ByVal NewData As String)
' Dateiname für die Verknüpfungsdatei (*.lnk).
  mFileName = NewData
End Property
Public Property Get FileName() As String
  FileName = mFileName
End Property
Public Property Let IconNumber(ByVal NewData As Long)
' Ressourcenposition des Icons in der durch IconPath
' repräsentierten Datei.
  mIconNumber = NewData
End Property
Public Property Get IconNumber() As Long
  IconNumber = mIconNumber
End Property
Public Property Let IconPath(ByVal NewData As String)
' Pfad zu einer Icondatei (*.ico) oder zu einer Datei,
' die Icons als Ressourcen enthält.
  mIconPath = NewData
End Property
Public Property Get IconPath() As String
  IconPath = mIconPath
End Property
Public Property Let Hotkey( _
  ByVal NewData As VBRUN.KeyCodeConstants)
' KeyCode der Taste, die im Zusammenhang mit den durch
' HotkeyModifiers beschriebenen Sondertasten als
' Tastenkombination verwendet werden soll.
  mHotkey = NewData
End Property
Public Property Get Hotkey() As VBRUN.KeyCodeConstants
  Hotkey = mHotkey
End Property
Public Property Let HotkeyModifiers( _
  ByVal NewData As VBRUN.ShiftConstants)
' Sondertasten (Strg, Shift und Alt), die zusammen mit der
' durch HotKey repräsentierten Taste als Tastenkombination
' verwendet werden.
  mHotkeyModifiers = NewData
End Property
Public Property Get HotkeyModifiers() As VBRUN.ShiftConstants
  HotkeyModifiers = mHotkeyModifiers
End Property
Public Property Let StartupWindowState( _
  ByVal NewData As StartupWindowStates)
' Festgelegter Öffnungszustand des Startfensters der Anwendung.
' Die Anwendung muss diese Möglichkeit der Einstellung
' unterstützen.
  mStartupWindowState = NewData
End Property
Public Property Get StartupWindowState() As StartupWindowStates
  StartupWindowState = mStartupWindowState
End Property
Public Property Let Arguments(ByVal NewData As String)
' Befehlszeilenargumente für den Aufruf der Anwendung.
  mArguments = NewData
End Property
Public Property Get Arguments() As String
  Arguments = mArguments
End Property
Public Property Let RelativePath(ByVal NewData As String)
' Relativer Pfad zur Anwendung.
  mRelativePath = NewData
End Property
Public Property Get RelativePath() As String
  RelativePath = mRelativePath
End Property
Public Property Let WorkingDirectory(ByVal NewData As String)
' Verzeichnis, das als Arbeitsverzeichnis verwendet werden soll.
  mWorkingDirectory = NewData
End Property
Public Property Get WorkingDirectory() As String
  WorkingDirectory = mWorkingDirectory
End Property
Public Property Let Path(ByVal NewData As String)
' Pfad zur verknüpften Anwendungsdatei.
  mPath = NewData
End Property
Public Property Get Path() As String
  Path = mPath
End Property
Public Property Let Description(ByVal NewData As String)
' Beschreibung der Verknüpfung
  mDescription = NewData
End Property
Public Property Get Description() As String
  Description = mDescription
End Property
' ---- CLASS INITIALIZE / TERMINATE ----
Private Sub Class_Initialize()
  mStartupWindowState = swsNormal ' Standardwert
End Sub
Private Sub Class_Terminate()
  ' Keine Aufräumarbeiten erforderlich.
End Sub
' ---- ÖFFENTLICHE FUNKTIONEN ----
Public Function LoadLink() As Boolean
' Lädt Informationen aus einer Verknüpfungsdatei und
' legt sie im cShellLink-Objekt ab.
' Rückgabewert: True bei Erfolg.
' Hinweis: Funktion löst im Bedarfsfall Fehler aus!
Const INFOTIPSIZE As Long = 1024
Const MAX_PATH As Long = 260
Dim PersistFile   As IShellLinkA.IPersistFile
Dim ShellLink   As IShellLinkA.ShellLinkA
Dim win32FindData   As IShellLinkA.WIN32_FIND_DATA
Dim strBuffer  As String
Dim lFlags  As Long
  ' Existiert die angegebene Datei?
  If LenB(Dir$(mFileName)) = 0 Then
 Err.Raise 53, "cShellLink.LoadLink" ' File not found
  End If
  ' Mithilfe von IPersistFile Informationen laden
  Set ShellLink = New IShellLinkA.ShellLinkA
  Set PersistFile = ShellLink
  PersistFile.Load StrConv(mFileName, vbUnicode), 0
  Set PersistFile = Nothing
  strBuffer = Space$(INFOTIPSIZE) ' Speicherreservierung
  With ShellLink
 ' Die Verknüpfungsbeschreibung ermitteln
 .GetDescription strBuffer, INFOTIPSIZE
 mDescription = NullTrim(strBuffer)
 ' Den Pfad zum Verknüpfungsziel ermitteln
 .GetPath strBuffer, MAX_PATH, win32FindData, lFlags
 mPath = NullTrim(strBuffer)
 ' Das Arbeitsverzeichnis ermitteln
 .GetWorkingDirectory strBuffer, INFOTIPSIZE
 mWorkingDirectory = NullTrim(strBuffer)
 ' Die Aufrufargumente ermitteln
 .GetArguments strBuffer, INFOTIPSIZE
 mArguments = NullTrim(strBuffer)
 ' Dateinamen und Icon-Nummer für das
 ' Verknüpfungsicon ermitteln
 .GetIconLocation strBuffer, INFOTIPSIZE, mIconNumber
 mIconPath = NullTrim(strBuffer)
 ' Tastenkombination ermitteln
 .GetHotkey mHotkey
 mHotkeyModifiers = mHotkey \ &H100& ' HiByte = Modifiers
 mHotkey = mHotkey And &HFF&   ' LoByte = Taste
 ' Initialen Fensterstil beim Ausführen ermitteln
 .GetShowCmd mStartupWindowState
  End With
  LoadLink = True
End Function
Public Function SaveLink( _
  Optional ByVal Overwrite As Boolean = True) As Boolean
' Speichert die Informationen des cShellLink-Objekts
' als Verknüpfung unter dem in FileName angegebenen Pfad ab.
' Setzen Sie Overwrite auf False, wenn Sie eine bereits
' bestehende Datei nicht überschreiben wollen.
' Rückgabewert: True bei Erfolg.
' Hinweis: Funktion löst im Bedarfsfall Fehler aus!
Dim PersistFile As IShellLinkA.IPersistFile
Dim ShellLink   As IShellLinkA.ShellLinkA
  ' Prüfen, ob die FileName-Eigenschaft belegt wurde
  If LenB(mFileName) = 0 Then
 Err.Raise 52, "cShellLink.SaveLink" ' Bad file name or number
  End If
  ' Prüfen, ob die in FileName angegebene Datei existiert
  If LenB(Dir$(mFileName, vbHidden)) > 0 Then
 If Overwrite = False Then
   Err.Raise 58, "cShellLink.SaveLink" ' File already exists
 End If
  End If
  Set ShellLink = New IShellLinkA.ShellLinkA
  ' IShellLink-Objekt mit Informationen versorgen
  With ShellLink
 .SetDescription mDescription
 .SetPath mPath
 .SetRelativePath mRelativePath, 0
 .SetWorkingDirectory mWorkingDirectory
 .SetArguments mArguments
 .SetHotkey CInt(mHotkey + &H100 * mHotkeyModifiers)
 .SetIconLocation mIconPath, mIconNumber
 .SetShowCmd mStartupWindowState
  End With
  ' Informationen mithilfe von IPersistFile abspeichern
  Set PersistFile = ShellLink
  PersistFile.Save StrConv(mFileName, vbUnicode), 0
  SaveLink = True
End Function
Public Function ModifiersDescription( _
  ByVal Modifiers As VBRUN.ShiftConstants) As String
' Gibt eine deutschsprachige Klartextbeschreibung für
' einen HotkeyModifiers-Wert zurück (z.B. "Strg + Alt")
  ' Ist eine Steuerungs-Taste mit im Spiel?
  If CBool(Modifiers And vbCtrlMask) Then
 ' "Strg" anfügen
 ModifiersDescription = "Strg"
  End If
  ' Ist eine Shift-Taste mit im Spiel?
  If CBool(Modifiers And vbShiftMask) Then
 ' Bei Bedarf " + " hinzufügen
 If LenB(ModifiersDescription) Then
   ModifiersDescription = ModifiersDescription & " + "
 End If
 ' "Umschalt" anfügen
 ModifiersDescription = ModifiersDescription & "Umschalt"
  End If
  ' Ist eine Alt-Taste mit im Spiel?
  If CBool(Modifiers And vbAltMask) Then
 ' Bei Bedarf " + " hinzufügen
 If LenB(ModifiersDescription) Then
   ModifiersDescription = ModifiersDescription & " + "
 End If
 ' "Alt" anfügen
 ModifiersDescription = ModifiersDescription & "Alt"
  End If
End Function
Public Function GetSystemFolderPath( _
 ByVal FolderID As ShellFolders) As String
' Ermittelt den Pfad zu bestimmten Systemordnern.
Dim ItemList As ITEMIDLIST
Dim sPath As String
  If SHGetSpecialFolderLocation(0, FolderID, ItemList) = 0 Then
 sPath = Space$(260)
 If SHGetPathFromIDList(ItemList.mkid.cb, sPath) <> 0 Then
   GetSystemFolderPath = Left$(sPath, _
 InStr(sPath, vbNullChar) - 1)
 End If
  End If
End Function
' ---- PRIVATE HILFSFUNKTIONEN ----
Private Function NullTrim( _
  ByRef StringToBeTrimmed As String) As String
' Gibt den Teil des übergebenen Strings zurück, der links vom
' ersten im String vorkommenden Nullzeichen (vbNullChar) steht.
' Kommt kein Nullzeichen vor, wird der übergebene String als
' Rückgabewert verwendet.
Dim lPos As Long  ' Verwendung als Position eines Zeichens
  ' Erstes vorkommendes Nullzeichen suchen:
  lPos = InStr(StringToBeTrimmed, vbNullChar)
  If lPos > 0 Then
 ' Teilstring vor dem Nullzeichen zurückgeben:
 NullTrim = Left$(StringToBeTrimmed, lPos - 1)
  Else
 ' Mangels Nullzeichen den kompletten String zurückgeben:
 NullTrim = StringToBeTrimmed
  End If
End Function


