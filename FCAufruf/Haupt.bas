Attribute VB_Name = "Haupt"
Public Const FcExePfad$ = "C:\Program Files (x86)\FreeCommander XE\FreeCommander.exe"
Public Const FcSetupExe$ = "\\linux1\daten\down\FreeCommanderXE-32_setup.exe"

Sub Main() ' ruft den FreeCommander mit P:\ links und Archiv-Unterverzeichnis eines Patienten rechts auf
 ' Analog zu ExpAufruf (dort Windows-Explorer), Aufruf des FreeCommander wie OeffneDualPane in NVerb (Haupt.bas)
 Const svz$ = "P:\dok\", asvz$ = "\\linux1\daten\Patientendokumente\dok\", uevz$ = "c:\gdt\", uedat$ = uevz & "turbcust.gdt"
 Dim zeile$, pid$, nachname$, uvz$, fcCmd$, fehlertext$
 On Error Resume Next
 MkDir uevz
 On Error GoTo fehler
 Open uedat For Input As #65
 Do While Not EOF(65)
  Line Input #65, zeile
  If zeile Like "???3000*" Then ' Patientennummer
   pid = Mid$(zeile, 8)
  ElseIf zeile Like "???3101*" Then ' Nachname (fuer den Namensfilter "Pat" im FreeCommander)
   nachname = Mid$(zeile, 8)
  End If
 Loop ' While Not EOF(65)
 Close #65
 If pid = "" Then
  MsgBox "In '" & uedat & "' wurde keine Patientennummer gefunden."
  Exit Sub
 End If
 uvz = svz & pid & "\"
 If Dir(uvz) = "" Then uvz = asvz & pid & "\"
 If Dir(uvz) = "" Then
  MsgBox "Archivordner '" & uvz & "' für Pat. " & pid & " (noch) nicht gefunden."
  Exit Sub
 End If
 Call FcEnsureInstalledAndConfigured
 Call FcSetPatFilterIncludeMask(nachname)
 ' rechter Pfad ohne abschliessenden Backslash, damit er nicht die schliessenden Anfuehrungszeichen maskiert
 fcCmd = Chr$(34) & FcExePfad & Chr$(34) & " /N /L=" & Chr$(34) & "P:\" & Chr$(34) & " /R=" & Chr$(34) & left$(uvz, Len(uvz) - 1) & Chr$(34)
 Call Shell(fcCmd, vbNormalFocus)
 Exit Sub
fehler:
 fehlertext = Err.Description
 Close #65
 MsgBox "Fehler beim Öffnen des Dateimanagers: " & fehlertext
End Sub ' Main()

' Ab hier unveraendert aus NVerb (Haupt.bas: FcEnsureInstalledAndConfigured bis FcSetPatFilterIncludeMask) uebernommen.
' Aenderungen dort bitte hier nachziehen.
' Installiert FreeCommander bei Bedarf still nach (Installer von \\linux1\daten\down\) und legt die
' Filter-/Farb-Vorlage in FreeCommander.ini an, falls sie fehlt. Wird bei jedem Alt+I-Aufruf geprueft -
' beides sind nur schnelle Dir$/ini-Existenzchecks, ausser beim allerersten Aufruf auf einem neuen PC
' (dort dauert die stille Installation einmalig einige Sekunden).
Sub FcEnsureInstalledAndConfigured()
 If Len(Dir$(FcExePfad)) = 0 Then
  If Len(Dir$(FcSetupExe)) = 0 Then Exit Sub ' Installer gerade nicht erreichbar
  Call Shell(Chr$(34) & FcSetupExe & Chr$(34) & " /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-", vbNormalFocus)
  Dim t0#
  t0 = Timer
  Do While Len(Dir$(FcExePfad)) = 0
   DoEvents
   If Timer - t0 > 120 Then Exit Sub ' Timeout, Installation offenbar fehlgeschlagen
  Loop
 End If
 Call FcEnsureFilterTemplate
 Call FcEnsureViewTemplate
End Sub ' FcEnsureInstalledAndConfigured

' FreeCommander.ini ist UTF-8 (nicht ANSI!) - deshalb Lesen/Schreiben ausschliesslich ueber
' ADODB.Stream, niemals ueber WritePrivateProfileStringA/GetPrivateProfileStringA (die schreiben
' im ANSI-Codepage des Rechners und zerstoeren dabei die UTF-8-Datei - FreeCommander-Fehlermeldung
' dann "No mapping for the Unicode character..." - siehe Vorfall 16.9.26.
Function FcReadIni$(pfad$)
 Dim st As New ADODB.Stream
 st.Type = adTypeText
 st.Charset = "utf-8"
 st.Open
 st.LoadFromFile pfad
 FcReadIni = st.ReadText
 st.Close
End Function ' FcReadIni$

Sub FcWriteIni(pfad$, inhalt$)
 Dim st As New ADODB.Stream
 st.Type = adTypeText
 st.Charset = "utf-8"
 st.Open
 st.WriteText inhalt
 st.SaveToFile pfad, adSaveCreateOverWrite
 st.Close
End Sub ' FcWriteIni

' Setzt (oder legt neu an) einen Schluessel innerhalb einer Section eines im Speicher gehaltenen
' ini-Texts. Reine Textbehandlung (kein Win32-Profile-API), damit die UTF-8-Kodierung erhalten bleibt.
Sub FcIniSetKey(ByRef inhalt$, section$, Schluessel$, wert$)
 Dim zeilen() As String
 zeilen = Split(inhalt, vbCrLf)
 Dim i&, secStart&, secEnd&
 secStart = -1
 For i = 0 To UBound(zeilen)
  If Trim$(zeilen(i)) = "[" & section & "]" Then
   secStart = i
   secEnd = UBound(zeilen)
   Dim j&
   For j = i + 1 To UBound(zeilen)
    If left$(Trim$(zeilen(j)), 1) = "[" Then
     secEnd = j - 1
     Exit For
    End If
   Next j
   Exit For
  End If
 Next i
 If secStart = -1 Then
  ReDim Preserve zeilen(UBound(zeilen) + 2)
  zeilen(UBound(zeilen) - 1) = "[" & section & "]"
  zeilen(UBound(zeilen)) = Schluessel & "=" & wert
  inhalt = Join(zeilen, vbCrLf)
  Exit Sub
 End If
 For i = secStart + 1 To secEnd
  If left$(Trim$(zeilen(i)), Len(Schluessel) + 1) = Schluessel & "=" Then
   zeilen(i) = Schluessel & "=" & wert
   inhalt = Join(zeilen, vbCrLf)
   Exit Sub
  End If
 Next i
 Dim neu() As String
 ReDim neu(UBound(zeilen) + 1)
 For i = 0 To secStart
  neu(i) = zeilen(i)
 Next i
 neu(secStart + 1) = Schluessel & "=" & wert
 For i = secStart + 1 To UBound(zeilen)
  neu(i + 1) = zeilen(i)
 Next i
 inhalt = Join(neu, vbCrLf)
End Sub ' FcIniSetKey

' Liest einen Schluesselwert aus einem im Speicher gehaltenen ini-Text (Gegenstueck zu FcIniSetKey).
Function FcIniGetKey$(inhalt$, section$, Schluessel$)
 Dim zeilen() As String
 zeilen = Split(inhalt, vbCrLf)
 Dim i&, secStart&, secEnd&
 secStart = -1
 For i = 0 To UBound(zeilen)
  If Trim$(zeilen(i)) = "[" & section & "]" Then
   secStart = i
   secEnd = UBound(zeilen)
   Dim j&
   For j = i + 1 To UBound(zeilen)
    If left$(Trim$(zeilen(j)), 1) = "[" Then
     secEnd = j - 1
     Exit For
    End If
   Next j
   Exit For
  End If
 Next i
 If secStart = -1 Then Exit Function
 For i = secStart + 1 To secEnd
  If left$(Trim$(zeilen(i)), Len(Schluessel) + 1) = Schluessel & "=" Then
   FcIniGetKey = Mid$(Trim$(zeilen(i)), Len(Schluessel) + 2)
   Exit Function
  End If
 Next i
End Function ' FcIniGetKey$

' Legt die beiden Filter (Pat/Scan) und ihre Farbzuordnung in FreeCommander.ini an, falls die Vorlage
' fehlt (erkannt am Scan-Filter, der "Br5_" enthalten muss). Ueberschreibt dabei bewusst auch evtl.
' individuell abweichende Filter unter denselben GUIDs - siehe Absprache vom 15.9.26.
Sub FcEnsureFilterTemplate()
 Dim iniPfad$
 iniPfad = Environ$("LOCALAPPDATA") & "\FreeCommanderXE\Settings\FreeCommander.ini"
 If Len(Dir$(iniPfad)) = 0 Then Exit Sub
 Dim inhalt$
 inhalt = FcReadIni(iniPfad)
 If InStr(inhalt, "Br5_") <> 0 Then Exit Sub ' Vorlage schon vorhanden
 Call FcIniSetKey(inhalt, "FcFilters", "1", "FE52C301-625F-43FF-AB48-DF7BB370B7FC")
 Call FcIniSetKey(inhalt, "FcFilters", "2", "FD6A2283-715D-4C3D-8856-488B6E71B568")
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "Title", "Pat")
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "UsingType", "0")
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "IncludeMask", "*")
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "ExcludeMask", "")
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "MaskAsRegExpr", "0")
 Call FcIniSetKey(inhalt, "Filter_FD6A2283-715D-4C3D-8856-488B6E71B568", "Title", "Scan")
 Call FcIniSetKey(inhalt, "Filter_FD6A2283-715D-4C3D-8856-488B6E71B568", "UsingType", "0")
 Call FcIniSetKey(inhalt, "Filter_FD6A2283-715D-4C3D-8856-488B6E71B568", "IncludeMask", "Br5_*;Br_*;Eps_*;")
 Call FcIniSetKey(inhalt, "Filter_FD6A2283-715D-4C3D-8856-488B6E71B568", "ExcludeMask", "")
 Call FcIniSetKey(inhalt, "Filter_FD6A2283-715D-4C3D-8856-488B6E71B568", "MaskAsRegExpr", "0")
 Call FcIniSetKey(inhalt, "ItemColorsByFileType", "1", "<Filter>:FE52C301-625F-43FF-AB48-DF7BB370B7FC|255")
 Call FcIniSetKey(inhalt, "ItemColorsByFileType", "2", "<Filter>:FD6A2283-715D-4C3D-8856-488B6E71B568|26367")
 Call FcWriteIni(iniPfad, inhalt)
End Sub ' FcEnsureFilterTemplate

' Legt Detailansicht (4 Spalten: Name/Groesse/Geaendert am/Attribute) und absteigende
' Sortierung nach Aenderungsdatum als Panel-Standard an, falls sie fehlt (erkannt an der
' Sektion FcDetailedViews_fc_default_view). Werte 1:1 von der Referenz-Installation
' uebernommen - siehe Absprache vom 16.9.26.
Sub FcEnsureViewTemplate()
 Dim iniPfad$
 iniPfad = Environ$("LOCALAPPDATA") & "\FreeCommanderXE\Settings\FreeCommander.ini"
 If Len(Dir$(iniPfad)) = 0 Then Exit Sub
 Dim inhalt$
 inhalt = FcReadIni(iniPfad)
 Call FcIniSetKey(inhalt, "MainPanel", "LeftViewStyle", "3")
 Call FcIniSetKey(inhalt, "MainPanel", "RightViewStyle", "3")
 Call FcIniSetKey(inhalt, "MainPanel", "LeftSortColumn", "0,9,0,2")
 Call FcIniSetKey(inhalt, "MainPanel", "RightSortColumn", "0,9,0,2")
 Call FcIniSetKey(inhalt, "MainPanel", "LeftDetailsProfile", "fc_default_view")
 Call FcIniSetKey(inhalt, "MainPanel", "RightDetailsProfile", "fc_default_view")
 Call FcIniSetKey(inhalt, "Form", "SortDirAlwaysOnEnd", "1") ' Dateien vor Ordnern
 Call FcIniSetKey(inhalt, "FcDetailedViews", "1", "fc_default_view")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "LoadShellTitle", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "Condition", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "ShowExtensionInCaption", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "ItemsCountRecursive", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "SizeCountRecursive", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "LastSorting", "0,9,0,2")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "AutoSizeNameColumn", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Name", "Name")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Format", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Align", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Width", "49")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Sort", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_ContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_ShellContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_RefValue", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_Content", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "1col_LvIndex", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Name", "Größe Automatisch")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Format", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Align", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Width", "9")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Sort", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_ContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_ShellContentType", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_RefValue", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_Content", "6")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "2col_LvIndex", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Name", "Geändert am")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Format", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Align", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Width", "22")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Sort", "1")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_ContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_ShellContentType", "2")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_RefValue", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_Content", "9")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "3col_LvIndex", "2")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Name", "Attribute")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Format", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Align", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Width", "6")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Sort", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_ContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_ShellContentType", "0")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_RefValue", "")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_Content", "7")
 Call FcIniSetKey(inhalt, "FcDetailedViews_fc_default_view", "4col_LvIndex", "3")
 Dim tabGuidL$, tabGuidR$
 tabGuidL = FcIniGetKey(inhalt, "TfcTabControl_Left", "ActiveTab")
 tabGuidR = FcIniGetKey(inhalt, "TfcTabControl_Right", "ActiveTab")
 If Len(tabGuidL) <> 0 Then
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidL, "ViewStyle", "3")
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidL, "Sort", "0,9,0,2")
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidL, "DetailedView", "fc_default_view")
 End If
 If Len(tabGuidR) <> 0 Then
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidR, "ViewStyle", "3")
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidR, "Sort", "0,9,0,2")
  Call FcIniSetKey(inhalt, "Tab_" & tabGuidR, "DetailedView", "fc_default_view")
 End If
 Call FcWriteIni(iniPfad, inhalt)
End Sub ' FcEnsureViewTemplate

' Setzt den IncludeMask-Wert des vorbereiteten Pat-Filters (Name enthaelt Nachname) in FreeCommander.ini.
' Der Filter selbst (Titel Pat, GUID FE52C301-625F-43FF-AB48-DF7BB370B7FC) sowie der zweite, statische
' Filter Scan (Br5_/Br_/Eps_) und die zugehoerigen Farbklassen muessen einmalig pro PC in FreeCommander
' selbst angelegt werden (Tools > Einstellungen > Filter definieren, dann Farbe nach Dateityp > vordefinierten
' Filter waehlen) - siehe Notiz vom 15.9.26.
Sub FcSetPatFilterIncludeMask(nachname$)
 If Len(nachname) = 0 Then Exit Sub
 Dim iniPfad$
 iniPfad = Environ$("LOCALAPPDATA") & "\FreeCommanderXE\Settings\FreeCommander.ini"
 If Len(Dir$(iniPfad)) = 0 Then Exit Sub ' FreeCommander wurde auf diesem PC noch nie gestartet/eingerichtet
 Dim inhalt$
 inhalt = FcReadIni(iniPfad)
 Call FcIniSetKey(inhalt, "Filter_FE52C301-625F-43FF-AB48-DF7BB370B7FC", "IncludeMask", "*" & nachname & "*")
 Call FcWriteIni(iniPfad, inhalt)
End Sub ' FcSetPatFilterIncludeMask
