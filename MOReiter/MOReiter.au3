#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=MOReiter.ico
#AutoIt3Wrapper_Outfile=MOReiter.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Kompilieren mit Icon: Aut2exe.exe /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico
; MOReiter: waehlt in Medical Office den Reiter "Kartei", "Krankenblatt", "ePA" oder "ePAAbr" per Mausklick
; Aufruf:  MOReiter.exe Kartei | Krankenblatt | ePA | ePAAbr  -> einmal klicken und beenden (z.B. aus VB6 per Shell)
;          MOReiter.exe Filter 1..24             -> Kartei und dort den n-ten Filter waehlen
;          MOReiter.exe Wechsel                  -> Kartei, bzw. Krankenblatt, wenn Kartei schon aktiv ist
;          MOReiter.exe ePAWechsel               -> ePA, bzw. ePAAbr, wenn ePA schon aktiv ist
;          MOReiter.exe                          -> bleibt resident mit Hotkeys:
;            Strg+Alt+K        Kartei, ist Kartei schon aktiv: Krankenblatt
;            Strg+Alt+L        Krankenblatt
;            Strg+Alt+P        ePA, ist ePA schon aktiv: ePAAbr
;            Strg+Alt+Umsch+K  Kalibrieren: Maus auf Reiter "Kartei" halten und druecken
;            Strg+Alt+Umsch+L  Kalibrieren: Maus auf Reiter "Krankenblatt" halten und druecken
;            Strg+Alt+Umsch+P  Kalibrieren: Maus auf Reiter "ePA" halten und druecken
;            Strg+Alt+Umsch+A  Kalibrieren: Maus auf Reiter "ePAAbr" halten und druecken
;            Strg+Alt+ ^ 1 2 3 4 5 6 7 8 9 0 sz Akut F2 .. F12
;                              Kartei (Reiter in [Filter] Reiter=) und dort der 1., 2., ... 24. Filter
;                              links ("Alle Eintraege", "Dateien", "RR Gewicht", ...)
;            Strg+Alt+Umsch+ dieselbe Taste  Kalibrieren: Maus auf diesen Filter halten und druecken;
;                              beim 1. Filter (^) wird die Position gemerkt, bei den anderen der Zeilenabstand
;            Mit AltGr (statt linker Strg+Alt) gedrueckt, werden die Tasten durchgereicht, so dass
;            AltGr+2 3 7 8 9 0 sz weiterhin hoch2 hoch3 { [ ] } \ liefern.
;            Beenden ueber das Tray-Menue (kein Strg+Alt+Q, das waere AltGr+Q = @)
; Laeuft Medical Office noch nicht, startet jeder Aufruf die Zentrale (medoff.exe) zum Anmelden.
; Die Reiterleiste ist ein Delphi-Control (TmoTabSet) ohne eigene Handles je Reiter,
; deshalb wird relativ zur linken oberen Ecke dieses Controls geklickt.
#include <Misc.au3>
#include <WinAPI.au3>
Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 4)

Const $MOFenster = "[REGEXPTITLE:^Medical Office; CLASS:OWL_Window]"
; INI im Benutzerprofil, weil das Programmverzeichnis unter "Program Files" nicht beschreibbar ist
Const $Ini = @AppDataDir & "\MOReiter\MOReiter.ini"

; Voreinstellungen (Pixel relativ zum Control), werden durch Kalibrieren in der INI ueberschrieben
Const $defCtrl = "[CLASS:TmoTabSet; INSTANCE:1]"
; (kalibriert am 27.9.26)
Const $defKarteiX = 53, $defKarteiY = 10, $defKrbX = 117, $defKrbY = 9
Const $defEpaX = 164, $defEpaY = 14, $defEpaAbrX = 212, $defEpaAbrY = 10
; Hintergrund des aktiven Reiters (hellblau, gemessen RGB 210,226,247), inaktive Reiter sind weiss
Const $defAktivFarbe = 0xD2E2F7, $defToleranz = 12
; Filterliste links in der Kartei: Mitte von "Alle Eintraege" relativ zur Reiterleiste und Zeilenabstand
; (geschaetzt aus einem Bildschirmfoto, genauer per Strg+Alt+Umsch+^ und z.B. Strg+Alt+Umsch+F12)
Const $defFilterReiter = "Kartei", $defFilterX = 46, $defFilterY = 83, $defFilterAbstand = 29
Const $defFilterWarten = 400 ; ms nach dem Reiterwechsel, bis die Filterliste steht
; Tasten fuer die Filter 1..24 in HotKeySet-Schreibweise und fuer Meldungen (sz und Akut als ChrW, Quelltext bleibt ASCII)
Global $FilterTasten[24] = ["{^}", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ChrW(223), ChrW(180), _
		"{F2}", "{F3}", "{F4}", "{F5}", "{F6}", "{F7}", "{F8}", "{F9}", "{F10}", "{F11}", "{F12}"]
Global $FilterNamen[24] = ["^", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ChrW(223), ChrW(180), _
		"F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]

If $CmdLine[0] > 1 And $CmdLine[1] = "Filter" Then
	Exit FilterWahl(Int($CmdLine[2]) - 1) ? 0 : 1
ElseIf $CmdLine[0] > 0 Then
	Exit Reiter($CmdLine[1]) ? 0 : 1
EndIf
If _Singleton("MOReiter_resident", 1) = 0 Then ; laeuft schon
	; sonst merkt man nicht, dass noch eine alte Version mit alten Hotkeys aktiv ist
;	MsgBox(64, "MOReiter", "MOReiter laeuft bereits." & @CRLF & "Fuer eine neue Version die alte zuerst ueber das Tray-Menue beenden.", 10)
	Exit
EndIf

; Icon aus der Datei daneben laden, falls es beim Kompilieren nicht in die exe gekommen ist
If FileExists(@ScriptDir & "\MOReiter.ico") Then TraySetIcon(@ScriptDir & "\MOReiter.ico")
TraySetToolTip("MOReiter: Strg+Alt+K Kartei/Wechsel, Strg+Alt+L Krankenblatt, Strg+Alt+P ePA/ePAAbr, Strg+Alt+^..F12 Filter")
Local $belegt = ""
If Not HotKeySet("^!k", "Wechsel") Then $belegt &= @CRLF & "Strg+Alt+K"
If Not HotKeySet("^!l", "Krankenblatt") Then $belegt &= @CRLF & "Strg+Alt+L"
If Not HotKeySet("^!p", "EpaWechsel") Then $belegt &= @CRLF & "Strg+Alt+P"
If Not HotKeySet("^!+k", "KalibKartei") Then $belegt &= @CRLF & "Strg+Alt+Umsch+K"
If Not HotKeySet("^!+l", "KalibKrankenblatt") Then $belegt &= @CRLF & "Strg+Alt+Umsch+L"
If Not HotKeySet("^!+p", "KalibEpa") Then $belegt &= @CRLF & "Strg+Alt+Umsch+P"
If Not HotKeySet("^!+a", "KalibEpaAbr") Then $belegt &= @CRLF & "Strg+Alt+Umsch+A"
; Filter-Hotkeys; was sich nicht registrieren laesst, wird nur gemeldet und bleibt weg
Local $fBelegt = "", $kBelegt = ""
For $i = 0 To UBound($FilterTasten) - 1
	If Not HotKeySet("^!" & $FilterTasten[$i], "Filter") Then $fBelegt &= " " & $FilterNamen[$i]
	If Not HotKeySet("^!+" & $FilterTasten[$i], "KalibFilter") Then $kBelegt &= " " & $FilterNamen[$i]
Next
If $fBelegt <> "" Then $belegt &= @CRLF & "Strg+Alt+" & $fBelegt
If $kBelegt <> "" Then $belegt &= @CRLF & "Strg+Alt+Umsch+" & $kBelegt
If $belegt <> "" Then MsgBox(48, "MOReiter", "Von einem anderen Programm belegt, wirkungslos:" & $belegt, 10)
While 1
	Sleep(100)
WEnd

Func Wechsel()
	Reiter("Wechsel")
EndFunc

Func Krankenblatt()
	Reiter("Krankenblatt")
EndFunc

Func EpaWechsel()
	Reiter("ePAWechsel")
EndFunc

Func KalibKartei()
	Kalibrieren("Kartei")
EndFunc

Func KalibKrankenblatt()
	Kalibrieren("Krankenblatt")
EndFunc

Func KalibEpa()
	Kalibrieren("ePA")
EndFunc

Func KalibEpaAbr()
	Kalibrieren("ePAAbr")
EndFunc

; Strg+Alt+Taste: Filter waehlen
Func Filter()
	If AltGr() Then Return Durchreichen("Filter")
	FilterWahl(FilterIndex(@HotKeyPressed))
EndFunc

; Strg+Alt+Umsch+Taste: Filter kalibrieren
Func KalibFilter()
	If AltGr() Then Return Durchreichen("KalibFilter")
	KalibrierFilter(FilterIndex(@HotKeyPressed))
EndFunc

; AltGr kommt bei Windows als linke Strg + rechte Alt an und loest daher dieselben Hotkeys aus
Func AltGr()
	Return _IsPressed("A5") ; rechte Alt-Taste
EndFunc

; schickt die Taste ohne den eigenen Hotkey weiter, damit z.B. AltGr+7 wieder { ergibt
Func Durchreichen($func)
	Local $hk = @HotKeyPressed
	HotKeySet($hk)
	Send($hk)
	HotKeySet($hk, $func)
EndFunc

; 0-basierte Nummer des Filters zur gedrueckten Hotkey-Taste, -1 falls unbekannt
Func FilterIndex($hk)
	$hk = StringRegExpReplace($hk, "^[\^!+]+", "")
	For $i = 0 To UBound($FilterTasten) - 1
		If $hk = $FilterTasten[$i] Then Return $i
	Next
	; "{^}" wird von @HotKeyPressed evtl. ohne Klammern geliefert
	If $hk = "" Or $hk = "^" Then Return 0
	Return -1
EndFunc

; wechselt bei Bedarf zur Kartei und klickt dort auf den Filter Nr. $i (0-basiert)
Func FilterWahl($i)
	If $i < 0 Or $i >= UBound($FilterTasten) Then Return Meldung("Unbekannter Filter")
	Local $hWnd, $hCtrl
	If Not HolLeiste($hWnd, $hCtrl) Then Return False
	Local $reiter = IniRead($Ini, "Filter", "Reiter", $defFilterReiter)
	If Not ReiterAktiv($reiter, $hWnd, $hCtrl) Then
		If Not Reiter($reiter) Then Return False
		Sleep(Int(IniRead($Ini, "Filter", "Warten", $defFilterWarten)))
	EndIf
	Local $x = Int(IniRead($Ini, "Filter", "X", $defFilterX))
	Local $y = Int(IniRead($Ini, "Filter", "Y", $defFilterY)) + $i * Number(IniRead($Ini, "Filter", "Abstand", $defFilterAbstand))
	Return KlickBei($hWnd, $hCtrl, $x, Round($y))
EndFunc

; klickt an die Stelle $x,$y relativ zur Reiterleiste, und zwar auf das Control, das dort liegt
; (die Filterliste ist ein anderes Control als die Reiterleiste)
Func KlickBei($hWnd, $hCtrl, $x, $y)
	Local $p = WinGetPos($hCtrl)
	If @error Then Return Meldung("Reiterleiste nicht gefunden")
	WinActivate($hWnd)
	WinWaitActive($hWnd, "", 2)
	If IniRead($Ini, "Allgemein", "EchteMaus", "0") = "1" Then
		Local $alt = MouseGetPos()
		Opt("MouseCoordMode", 1)
		MouseClick("left", $p[0] + $x, $p[1] + $y, 1, 0)
		MouseMove($alt[0], $alt[1], 0)
		Return True
	EndIf
	Local $pt = DllStructCreate("int X;int Y")
	$pt.X = $p[0] + $x
	$pt.Y = $p[1] + $y
	Local $hZiel = _WinAPI_WindowFromPoint($pt)
	; 2 = GA_ROOT: das Control muss zum Medical-Office-Fenster gehoeren
	If $hZiel = 0 Or _WinAPI_GetAncestor($hZiel, 2) <> $hWnd Then Return Meldung("Filterliste verdeckt oder nicht sichtbar")
	_WinAPI_ScreenToClient($hZiel, $pt)
	If Not ControlClick($hWnd, "", $hZiel, "left", 1, $pt.X, $pt.Y) Then Return Meldung("Klick fehlgeschlagen")
	Return True
EndFunc

; merkt sich beim 1. Filter die Position, bei jedem anderen den Zeilenabstand zum 1.
Func KalibrierFilter($i)
	If $i < 0 Then Return Meldung("Unbekannter Filter")
	Local $hWnd, $hCtrl
	If Not HolLeiste($hWnd, $hCtrl) Then Return False
	Opt("MouseCoordMode", 1)
	Local $m = MouseGetPos(), $p = WinGetPos($hCtrl)
	Local $x = $m[0] - $p[0], $y = $m[1] - $p[1]
	DirCreate(@AppDataDir & "\MOReiter")
	If $i = 0 Then
		IniWrite($Ini, "Filter", "X", $x)
		IniWrite($Ini, "Filter", "Y", $y)
		Hinweis("Filter 1 kalibriert: " & $x & ", " & $y)
	Else
		Local $abst = Round(($y - Int(IniRead($Ini, "Filter", "Y", $defFilterY))) / $i, 2)
		If $abst < 5 Then Return Meldung("Erst mit Strg+Alt+Umsch+^ den 1. Filter kalibrieren")
		IniWrite($Ini, "Filter", "Abstand", $abst)
		Hinweis("Filter-Zeilenabstand kalibriert: " & $abst)
	EndIf
	Return True
EndFunc

; sucht Medical Office und die sichtbare Reiterleiste
Func HolLeiste(ByRef $hWnd, ByRef $hCtrl)
	$hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	$hCtrl = ControlGetHandle($hWnd, "", IniRead($Ini, "Allgemein", "Control", $defCtrl))
	If @error Or Not BitAND(WinGetState($hCtrl), 2) Then Return Meldung("Reiterleiste nicht sichtbar (Patient geoeffnet?)")
	Return True
EndFunc

; liefert True bei Erfolg
Func Reiter($name)
	Local $x, $y
	Local $wechsel = ($name = "Wechsel" Or $name = "ePAWechsel")
	If Not $wechsel And Not HolPos($name, $x, $y) Then Return Meldung("Unbekannter Reiter: " & $name)
	Local $hWnd, $hCtrl
	If Not HolLeiste($hWnd, $hCtrl) Then Return False
	If $name = "Wechsel" Then
		$name = ReiterAktiv("Kartei", $hWnd, $hCtrl) ? "Krankenblatt" : "Kartei"
		HolPos($name, $x, $y)
	ElseIf $name = "ePAWechsel" Then
		$name = ReiterAktiv("ePA", $hWnd, $hCtrl) ? "ePAAbr" : "ePA"
		HolPos($name, $x, $y)
	EndIf

	If IniRead($Ini, "Allgemein", "EchteMaus", "0") = "1" Then
		; Ausweichweg, falls das Control auf gepostete Mausnachrichten nicht reagiert
		Local $p = WinGetPos($hCtrl), $alt = MouseGetPos()
		WinActivate($hWnd)
		WinWaitActive($hWnd, "", 2)
		Opt("MouseCoordMode", 1)
		MouseClick("left", $p[0] + $x, $p[1] + $y, 1, 0)
		MouseMove($alt[0], $alt[1], 0)
	Else
		WinActivate($hWnd)
		If Not ControlClick($hWnd, "", $hCtrl, "left", 1, $x, $y) Then Return Meldung("Klick fehlgeschlagen")
	EndIf
	Return True
EndFunc

Func HolPos($name, ByRef $x, ByRef $y)
	Switch $name
		Case "Kartei"
			$x = Int(IniRead($Ini, "Kartei", "X", $defKarteiX))
			$y = Int(IniRead($Ini, "Kartei", "Y", $defKarteiY))
		Case "Krankenblatt"
			$x = Int(IniRead($Ini, "Krankenblatt", "X", $defKrbX))
			$y = Int(IniRead($Ini, "Krankenblatt", "Y", $defKrbY))
		Case "ePA"
			$x = Int(IniRead($Ini, "ePA", "X", $defEpaX))
			$y = Int(IniRead($Ini, "ePA", "Y", $defEpaY))
		Case "ePAAbr"
			$x = Int(IniRead($Ini, "ePAAbr", "X", $defEpaAbrX))
			$y = Int(IniRead($Ini, "ePAAbr", "Y", $defEpaAbrY))
		Case Else
			Return False
	EndSwitch
	Return True
EndFunc

; prueft am Bildschirm, ob der Reiter $name hellblau hinterlegt, also aktiv ist
; (eine PixelSearch ueber ein kleines Rechteck dauert nur wenige Millisekunden)
Func ReiterAktiv($name, $hWnd, $hCtrl)
	Local $x, $y
	HolPos($name, $x, $y)
	If Not WinActive($hWnd) Then
		WinActivate($hWnd)
		If Not WinWaitActive($hWnd, "", 2) Then Return False
		Sleep(100) ; bis das Fenster neu gezeichnet ist
	EndIf
	Local $p = WinGetPos($hCtrl)
	If @error Then Return False
	Local $farbe = Int(IniRead($Ini, "Allgemein", "AktivFarbe", $defAktivFarbe))
	Local $tol = Int(IniRead($Ini, "Allgemein", "Toleranz", $defToleranz))
	Opt("PixelCoordMode", 1)
	PixelSearch($p[0] + $x - 12, $p[1] + $y - 5, $p[0] + $x + 12, $p[1] + $y + 5, $farbe, $tol)
	Return Not @error
EndFunc

; merkt sich die aktuelle Mausposition relativ zur Reiterleiste
Func Kalibrieren($name)
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $hCtrl = ControlGetHandle($hWnd, "", IniRead($Ini, "Allgemein", "Control", $defCtrl))
	If @error Then Return Meldung("Reiterleiste nicht gefunden")
	Opt("MouseCoordMode", 1)
	Local $m = MouseGetPos(), $p = WinGetPos($hCtrl)
	Local $x = $m[0] - $p[0], $y = $m[1] - $p[1]
	If $x < 0 Or $y < 0 Or $x >= $p[2] Or $y >= $p[3] Then Return Meldung("Maus steht nicht auf der Reiterleiste")
	DirCreate(@AppDataDir & "\MOReiter")
	IniWrite($Ini, $name, "X", $x)
	IniWrite($Ini, $name, "Y", $y)
	Hinweis($name & " kalibriert: " & $x & ", " & $y)
	Return True
EndFunc

; startet die Zentrale, falls medoff.exe noch gar nicht laeuft (sonst ist z.B. nur der Login offen)
; liefert immer False, weil der Reiter erst nach dem Anmelden gewaehlt werden kann
Func MOStarten()
	If ProcessExists("medoff.exe") Then Return Meldung("Medical Office nicht angemeldet")
	Local $pfad, $pfade[4] = ["C:\medoff\medoff.exe", "C:\INDAMED\medoff.exe", "D:\medoff\medoff.exe", "D:\INDAMED\medoff.exe"]
	For $pfad In $pfade
		If FileExists($pfad) Then
			Run($pfad, StringLeft($pfad, StringInStr($pfad, "\", 0, -1) - 1))
			Return Meldung("Medical Office wird gestartet, bitte anmelden")
		EndIf
	Next
	Return Meldung("Medical Office nicht gefunden")
EndFunc

Func Meldung($txt)
	If $CmdLine[0] > 0 Then
		ConsoleWriteError($txt & @CRLF)
	Else
		Hinweis($txt)
	EndIf
	Return False
EndFunc

; ToolTip an der Maus, weil TrayTip unter Windows 10/11 eine Benachrichtigung ist,
; die bei abgeschalteten Benachrichtigungen oder "Nicht stoeren" gar nicht erscheint
Func Hinweis($txt)
	ToolTip($txt, Default, Default, "MOReiter")
	AdlibRegister("HinweisWeg", 3000)
EndFunc

Func HinweisWeg()
	AdlibUnRegister("HinweisWeg")
	ToolTip("")
EndFunc
