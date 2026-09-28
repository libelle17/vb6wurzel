#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=MOReiter.ico
#AutoIt3Wrapper_Outfile=MOReiter.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Kompilieren mit Icon: Aut2exe.exe /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico
; MOReiter: waehlt in Medical Office den Reiter "Kartei", "Krankenblatt", "ePA" oder "ePAAbr" per Mausklick
; Aufruf:  MOReiter.exe Kartei | Krankenblatt | ePA | ePAAbr  -> einmal klicken und beenden (z.B. aus VB6 per Shell)
;          MOReiter.exe LetztePatienten | Menue  -> Pfeil neben der Patientensuche bzw. drei Striche links oben
;          MOReiter.exe Filter 1..24             -> Kartei und dort den n-ten Filter waehlen
;          MOReiter.exe Wechsel                  -> Kartei, bzw. Krankenblatt, wenn Kartei schon aktiv ist
;          MOReiter.exe ePAWechsel               -> ePA, bzw. ePAAbr, wenn ePA schon aktiv ist
;          MOReiter.exe                          -> bleibt resident mit Hotkeys:
;            Strg+Alt+K        Kartei, ist Kartei schon aktiv: Krankenblatt
;            Strg+Alt+L        Krankenblatt
;            Strg+Alt+P        ePA, ist ePA schon aktiv: ePAAbr
;            Strg+Alt+Z        Liste der zuletzt geoeffneten Patienten (Pfeil rechts neben der Patientensuche)
;            Strg+Alt+Leertaste  Hauptmenue (drei Striche links oben)
;            Strg+Alt+H        in der markierten Zeile der Krankenblatt-Liste den hellblauen Pfeil nach
;                              oben anklicken (in die ePA hochladen), im folgenden Fenster "Hochladen" druecken
;                              und einen danach erscheinenden Leistungsdialog mit "Uebernehmen" bestaetigen
;            Strg+Alt+A        nur in der MO-Tagesuebersicht bei "Offene To-Dos": markierte Zeile als
;                              erledigt abhaken und die nachrueckende Zeile markieren
;            Strg+Alt+Umsch+A  in der Tagesuebersicht: Maus auf die Spalte "Bereich" halten und druecken
;            Strg+Alt+Umsch+K  Kalibrieren: Maus auf Reiter "Kartei" halten und druecken
;            Strg+Alt+Umsch+L  Kalibrieren: Maus auf Reiter "Krankenblatt" halten und druecken
;            Strg+Alt+Umsch+P  Kalibrieren: Maus auf Reiter "ePA" halten und druecken
;            Strg+Alt+Umsch+A  Kalibrieren: Maus auf Reiter "ePAAbr" halten und druecken
;            Strg+Alt+ ^ 1 2 3 4 5 6 7 8 9 0 sz Akut F2 .. F12
;                              Kartei (Reiter in [Filter] Reiter=) und dort der 1., 2., ... 24. Filter
;                              links ("Alle Eintraege", "Dateien", "RR Gewicht", ...)
;            Strg+Alt+Umsch+ dieselbe Taste  Kalibrieren: Maus auf diesen Filter halten und druecken;
;                              beim 1. Filter (^) wird die Position gemerkt, bei den anderen der Zeilenabstand
;            Die Tasten werden per Tastatur-Hook abgefangen (nicht per HotKeySet), weil nur so linkes Alt
;            von AltGr unterschieden werden kann: AltGr kommt als linke Strg + rechte Alt an und bleibt
;            unberuehrt, so dass AltGr+2 3 7 8 9 0 sz weiterhin hoch2 hoch3 { [ ] } \ liefern.
;            (Das fruehere Durchreichen per Send liess gelegentlich die Strg-Taste haengen.)
;            Beenden ueber das Tray-Menue (kein Strg+Alt+Q, das waere AltGr+Q = @)
; Laeuft Medical Office noch nicht, startet jeder Aufruf die Zentrale (medoff.exe) zum Anmelden.
; Die Reiterleiste ist ein Delphi-Control (TmoTabSet) ohne eigene Handles je Reiter,
; deshalb wird relativ zur linken oberen Ecke dieses Controls geklickt.
#include <Misc.au3>
#include <WinAPI.au3>
#include <WindowsNotifsConstants.au3>
#include <ScreenCapture.au3>
Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 4)
; keine 250 ms Pause nach jedem Fensterbefehl; wo ein Fenster Zeit braucht, wird ausdruecklich gewartet
Opt("WinWaitDelay", 0)

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
; Tagesuebersicht (gemessen 28.9.26, Pixel relativ zur Liste): Spalten "Bereich" und "Name",
; farbige Randspalte links, Kopfzeile, Zeilenhoehe, Farbe des eingedrueckten Knopfs "Offene To-Dos"
Const $TagesKlasse = "TfrmTagesUebersicht"
Const $defBereichX = 470, $defNameX = 150, $defAbhakenWarten = 500
Const $RandSpalteX = 10, $KopfHoehe = 22, $ZeilenHoehe = 20, $OffenFarbe = 0x7ABEE7
; Krankenblatt-Liste (gemessen 28.9.26): Hintergrund der markierten Zeile, Kopfzeilenhoehe, Hoehe des
; Pfeilschafts ab Zeilenoberkante; Pfeil nach oben hellblau = noch nicht in der ePA, dunkel = schon drin
Const $KbMarkiertFarbe = 0xD3E3F7, $KbKopfHoehe = 18, $KbPfeilY = 11
Const $KbPfeilHell = 0x6CC1EF, $KbPfeilDunkel = 0x233658, $KbPfeilArm = 0x809FBA
; Fenster nach dem Pfeilklick, darin wird der Knopf "Hochladen" nach $defHochladenWarten ms gedrueckt
; (100 ms reichten nur, solange nach jedem Fensterbefehl noch 250 ms WinWaitDelay dazukamen)
Const $EpaDialog = "[REGEXPTITLE:^MEDICAL OFFICE - ePA; CLASS:TClientWindowForm]", $defHochladenWarten = 300
; danach evtl. Dialog zur Leistungsdokumentation mit "Uebernehmen", wird bis zu $defLeistungWarten ms erwartet
Const $LeistungDialog = "[TITLE:Medical Office; CLASS:TfrmDialogContainer]", $defLeistungWarten = 1500
; virtuelle Tastencodes fuer die Filter 1..24 (deutsches Layout: ^ = OEM_5, sz = OEM_4, Akut = OEM_6)
Global $FilterTasten[24] = [0xDC, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30, 0xDB, 0xDD, _
		0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x7B]
; Tastatur-Hook: der Rueckruf merkt sich nur die Aufgabe, ausgefuehrt wird sie in der Hauptschleife,
; weil Windows einen Hook, der zu lange braucht, stillschweigend abhaengt
Global $gHook = 0, $gAufgabe = "", $gGehalten = 0, $gGehaltenZeit = 0
; eingelesener Bildschirmausschnitt fuer die Pfeilpruefung (BildLesen/BildFarbe)
Global $gBild = 0, $gBildX = 0, $gBildY = 0, $gBildB = 0

If $CmdLine[0] > 1 And $CmdLine[1] = "Filter" Then
	Exit FilterWahl(Int($CmdLine[2]) - 1) ? 0 : 1
ElseIf $CmdLine[0] > 0 Then
	Exit Ausfuehren($CmdLine[1]) ? 0 : 1
EndIf
If _Singleton("MOReiter_resident", 1) = 0 Then ; laeuft schon
	; sonst merkt man nicht, dass noch eine alte Version mit alten Hotkeys aktiv ist
;	MsgBox(64, "MOReiter", "MOReiter laeuft bereits." & @CRLF & "Fuer eine neue Version die alte zuerst ueber das Tray-Menue beenden.", 10)
	Exit
EndIf

; Icon aus der Datei daneben laden, falls es beim Kompilieren nicht in die exe gekommen ist
If FileExists(@ScriptDir & "\MOReiter.ico") Then TraySetIcon(@ScriptDir & "\MOReiter.ico")
TraySetToolTip("MOReiter: Strg+Alt+K Kartei/Wechsel, L Krankenblatt, P ePA/ePAAbr, Z letzte Patienten, Leertaste Menue, ^..F12 Filter")
; sonst haelt ein Klick aufs Tray-Symbol das Skript an, und Windows haengt den unbeantworteten Hook ab
Opt("TrayAutoPause", 0)
Global $gRueckruf = DllCallbackRegister("TastenHook", "lresult", "int;wparam;lparam")
If Not HookErneuern() Then
	MsgBox(16, "MOReiter", "Tastatur-Hook konnte nicht eingerichtet werden.", 10)
	Exit 1
EndIf
OnAutoItExitRegister("HookEntfernen")
Local $erneuert = TimerInit()
; kurzes Sleep, damit der Hook die Tastatur nicht spuerbar verzoegert
While 1
	Sleep(10)
	If $gAufgabe <> "" Then
		Local $aufgabe = $gAufgabe
		$gAufgabe = ""
		Ausfuehren($aufgabe)
	EndIf
	; Windows entfernt einen Hook stillschweigend, wenn er einmal zu langsam antwortet (z.B. bei hoher
	; Last, Sperren, Aufwachen), das Programm liefe dann taub weiter; daher regelmaessig neu einhaengen
	If TimerDiff($erneuert) > 5000 Then
		HookErneuern()
		$erneuert = TimerInit()
	EndIf
WEnd

; haengt den neuen Hook ein, bevor der alte entfernt wird, damit keine Taste verloren geht
Func HookErneuern()
	Local $neu = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($gRueckruf), _WinAPI_GetModuleHandle(0))
	If $neu = 0 Then Return False
	Local $alt = $gHook
	$gHook = $neu
	If $alt Then _WinAPI_UnhookWindowsHookEx($alt)
	Return True
EndFunc

Func HookEntfernen()
	If $gHook Then _WinAPI_UnhookWindowsHookEx($gHook)
EndFunc

; faengt linke Alt + Strg [+ Umsch] + eigene Taste ab; alles andere, auch AltGr, geht unveraendert weiter
Func TastenHook($nCode, $wParam, $lParam)
	If $nCode >= 0 Then
		Local $kb = DllStructCreate($tagKBDLLHOOKSTRUCT, $lParam)
		Local $vk = $kb.vkCode
		If $wParam = $WM_KEYUP Or $wParam = $WM_SYSKEYUP Then
			; zum geschluckten Druecken auch das Loslassen schlucken
			If $vk = $gGehalten Then
				$gGehalten = 0
				Return 1
			EndIf
		Else
			; auch kuenstlich erzeugte Tasten (LLKHF_INJECTED) auswerten: RemotePC, TeamViewer u.ae.
			; liefern alle Tasten so; MOReiter selbst sendet nie Strg+Alt+Taste
			Local $aufgabe = HookAufgabe($vk)
			If $aufgabe <> "" Then
				; automatische Wiederholung beim Festhalten schlucken; nur innerhalb 1 s, weil
				; Fernsteuerungen das Loslassen manchmal verschlucken und die Taste sonst taub bliebe
				If $vk = $gGehalten And TimerDiff($gGehaltenZeit) < 1000 Then
					$gGehaltenZeit = TimerInit()
					Return 1
				EndIf
				$gGehalten = $vk
				$gGehaltenZeit = TimerInit()
				$gAufgabe = $aufgabe
				Return 1
			EndIf
		EndIf
	EndIf
	Return _WinAPI_CallNextHookEx($gHook, $nCode, $wParam, $lParam)
EndFunc

; Aufgabe zur Taste $vk, falls gerade Strg + linke Alt (ohne AltGr) gehalten werden, sonst ""
Func HookAufgabe($vk)
	If Not (Gedrueckt(0xA2) Or Gedrueckt(0xA3)) Or Not Gedrueckt(0xA4) Or Gedrueckt(0xA5) Then Return ""
	Local $kalib = Gedrueckt(0x10)
	Switch $vk
		Case 0x4B ; K
			Return $kalib ? "Kalib:Kartei" : "Wechsel"
		Case 0x4C ; L
			Return $kalib ? "Kalib:Krankenblatt" : "Krankenblatt"
		Case 0x50 ; P
			Return $kalib ? "Kalib:ePA" : "ePAWechsel"
		Case 0x41 ; A, in der Tagesuebersicht: abhaken bzw. Spalte "Bereich" kalibrieren
			If TagesAktiv() Then Return $kalib ? "KalibBereich" : "Abhaken"
			Return $kalib ? "Kalib:ePAAbr" : ""
		Case 0x48 ; H, nur in Medical Office
			Return ($kalib Or Not MOAktiv()) ? "" : "Hochladen"
		Case 0x5A ; Z
			Return $kalib ? "" : "LetztePatienten"
		Case 0x20 ; Leertaste
			Return $kalib ? "" : "Menue"
	EndSwitch
	For $i = 0 To UBound($FilterTasten) - 1
		If $vk = $FilterTasten[$i] Then Return ($kalib ? "KalibFilter:" : "Filter:") & $i
	Next
	Return ""
EndFunc

Func Gedrueckt($vk)
	Local $r = DllCall("user32.dll", "short", "GetAsyncKeyState", "int", $vk)
	Return Not @error And BitAND($r[0], 0x8000) <> 0
EndFunc

Func Ausfuehren($aufgabe)
	If $aufgabe = "LetztePatienten" Then Return LetztePatienten()
	If $aufgabe = "Menue" Then Return Menue()
	If $aufgabe = "Abhaken" Then Return Abhaken()
	If $aufgabe = "Hochladen" Then Return Hochladen()
	If $aufgabe = "KalibBereich" Then Return KalibBereich()
	Local $teil = StringSplit($aufgabe, ":", 2)
	If UBound($teil) = 1 Then Return Reiter($aufgabe)
	Switch $teil[0]
		Case "Kalib"
			Kalibrieren($teil[1])
		Case "Filter"
			FilterWahl(Int($teil[1]))
		Case "KalibFilter"
			KalibrierFilter(Int($teil[1]))
	EndSwitch
EndFunc

; ist das Hauptfenster von Medical Office aktiv? (im Hook, muss schnell sein)
Func MOAktiv()
	Local $h = _WinAPI_GetForegroundWindow()
	Return _WinAPI_GetClassName($h) = "OWL_Window" And StringLeft(_WinAPI_GetWindowText($h), 14) = "Medical Office"
EndFunc

; Krankenblatt-Liste (Container in Kartei, Krankenblatt, ePA): in der markierten Zeile auf den Pfeil
; nach oben klicken, der hellblau ist, solange der Eintrag noch nicht in die ePA hochgeladen ist.
; Die Liste (TmoStringGrid) hat keine ansprechbaren Zeilen, daher per Bildschirm: markierte Zeile an
; ihrem Hintergrund suchen, in ihrer ersten Textzeile nach einem hellblauen senkrechten Pfeilschaft
; tasten und zur Sicherheit die beiden schraegen Arme der Pfeilspitze pruefen (Text kann durch
; Kantenglaettung einzelne hellblaue Pixel haben, aber keine Pfeilspitze).
Func Hochladen()
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Or Not WinActive($hWnd) Then Return False
	Local $hGrid = KarteiListe($hWnd)
	If Not $hGrid Then Return Meldung("Keine Krankenblatt-Liste gefunden")
	If Not ModifierLos() Then Return Meldung("Strg und Alt bitte loslassen")
	Local $g = WinGetPos($hGrid)
	If @error Then Return False
	Opt("PixelCoordMode", 1)
	Opt("MouseCoordMode", 1)
	Local $p = PixelSearch($g[0] + 4, $g[1] + $KbKopfHoehe, $g[0] + 4, $g[1] + $g[3] - 3, $KbMarkiertFarbe, 6)
	If @error Then Return Meldung("Keine markierte Zeile sichtbar")
	Local $y = $p[1] + $KbPfeilY
	; den Streifen um die Suchhoehe einmal einlesen, statt jedes Pixel einzeln (je ca. 15-20 ms) abzufragen
	If Not BildLesen($g[0] + 2, $y - 24, $g[0] + $g[2] - 18, $y + 6) Then Return Meldung("Bildschirm nicht lesbar")
	Local $x = PfeilSchaft($g[0] + 4, $g[0] + $g[2] - 20, $y, $KbPfeilHell, True)
	Local $dunkel = ($x < 0) ? PfeilSchaft($g[0] + 4, $g[0] + $g[2] - 20, $y, $KbPfeilDunkel, False) : -1
	$gBild = 0
	If $x < 0 Then
		If $dunkel >= 0 Then Return Meldung("Schon in der ePA (Pfeil ist schwarz)")
		Return Meldung("Kein Pfeil zum Hochladen in der markierten Zeile")
	EndIf
	Local $alt = MouseGetPos()
	MouseClick("left", $x, $y, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	; im Fenster "MEDICAL OFFICE - ePA3.0 - Datensatz" kurz nach dem Erscheinen "Hochladen" druecken
	Local $hDlg = WinWait($EpaDialog, "", 15)
	If Not $hDlg Then Return Meldung("Fenster ""MEDICAL OFFICE - ePA"" ist nicht erschienen")
	Local $hKnopf = KnopfBereit($hDlg, "Hochladen", 5000)
	If Not $hKnopf Then Return Meldung("Knopf ""Hochladen"" nicht bereit")
	Sleep(Int(IniRead($Ini, "Hochladen", "Warten", $defHochladenWarten)))
	; schon offene Dialoge merken, damit nur ein neu erscheinender Leistungsdialog bestaetigt wird
	Local $vorher = WinList($LeistungDialog)
	For $i = 1 To $vorher[0][0]
		$vorher[$i][0] = BitAND(WinGetState($vorher[$i][1]), 2) ? "sichtbar" : ""
	Next
	If Not KnopfKlick($hDlg, $hKnopf) Then Return False
	; manchmal folgt nach dem Hochladen ein Dialog zur Leistungsdokumentation: dort "Uebernehmen"
	If Not WinWaitClose($hDlg, "", 30) Then Return True
	Local $t = TimerInit(), $max = Int(IniRead($Ini, "Hochladen", "LeistungWarten", $defLeistungWarten))
	While TimerDiff($t) < $max
		Local $liste = WinList($LeistungDialog)
		For $i = 1 To $liste[0][0]
			If Not BitAND(WinGetState($liste[$i][1]), 2) Or InWinList($vorher, $liste[$i][1]) Then ContinueLoop
			Local $hUeb = KnopfBereit($liste[$i][1], ChrW(220) & "bernehmen", 1000)
			If $hUeb Then
				Sleep(Int(IniRead($Ini, "Hochladen", "Warten", $defHochladenWarten)))
				Return KnopfKlick($liste[$i][1], $hUeb)
			EndIf
		Next
		Sleep(50)
	WEnd
	Return True
EndFunc

; wartet bis zu $ms, bis das Fenster $hDlg sichtbar und sein Knopf $text sichtbar und bedienbar ist
; (Delphi legt Fenster erst unsichtbar an, ein Klick zu frueh geht ins Leere); liefert den Knopf oder 0
Func KnopfBereit($hDlg, $text, $ms)
	Local $t = TimerInit()
	While TimerDiff($t) < $ms
		Local $h = ControlGetHandle($hDlg, "", "[CLASS:TmoButtonFlat; TEXT:" & $text & "]")
		If Not @error And BitAND(WinGetState($hDlg), 2) And BitAND(WinGetState($h), 2) _
				And ControlCommand($hDlg, "", $h, "IsEnabled", "") Then Return $h
		Sleep(20)
	WEnd
	Return 0
EndFunc

; nur sichtbare zaehlen: ein versteckt vorgehaltener Dialog, der spaeter erscheint, gilt als neu
Func InWinList($liste, $h)
	For $i = 1 To $liste[0][0]
		If $liste[$i][1] = $h And $liste[$i][0] = "sichtbar" Then Return True
	Next
	Return False
EndFunc

; die Liste im Krankenblatt-Container: bevorzugt die mit dem Fokus, sonst die im TKarteikarteForm
Func KarteiListe($hWnd)
	Local $h = ControlGetHandle($hWnd, "", ControlGetFocus($hWnd))
	If Not @error And _WinAPI_GetClassName($h) = "TmoStringGrid" _
			And _WinAPI_GetClassName(_WinAPI_GetParent($h)) = "TfrmKarteikarteControl" Then Return $h
	Local $hForm = ControlGetHandle($hWnd, "", "[CLASS:TKarteikarteForm; INSTANCE:1]")
	If @error Or Not BitAND(WinGetState($hForm), 2) Then Return 0
	Local $liste = _WinAPI_EnumChildWindows($hForm)
	If @error Then Return 0
	For $i = 1 To $liste[0][0]
		If $liste[$i][1] = "TmoStringGrid" Then Return $liste[$i][0]
	Next
	Return 0
EndFunc

; liest das Bildschirmrechteck $x1,$y1 - $x2,$y2 in den Speicher ($gBild, 32 Bit je Pixel, zeilenweise
; von oben); danach liefert BildFarbe() Pixel daraus
Func BildLesen($x1, $y1, $x2, $y2)
	Local $hBmp = _ScreenCapture_Capture("", $x1, $y1, $x2, $y2, False)
	If @error Or Not $hBmp Then Return False
	$gBildX = $x1
	$gBildY = $y1
	$gBildB = $x2 - $x1 + 1
	Local $h = $y2 - $y1 + 1
	$gBild = DllStructCreate("dword[" & $gBildB * $h & "]")
	Local $bmi = DllStructCreate("dword biSize;long biWidth;long biHeight;word biPlanes;word biBitCount;dword biCompression;dword biSizeImage;long biXPelsPerMeter;long biYPelsPerMeter;dword biClrUsed;dword biClrImportant")
	$bmi.biSize = DllStructGetSize($bmi)
	$bmi.biWidth = $gBildB
	$bmi.biHeight = -$h ; negativ: oberste Zeile zuerst
	$bmi.biPlanes = 1
	$bmi.biBitCount = 32
	Local $hDC = _WinAPI_GetDC(0)
	Local $r = DllCall("gdi32.dll", "int", "GetDIBits", "handle", $hDC, "handle", $hBmp, "uint", 0, "uint", $h, _
			"struct*", $gBild, "struct*", $bmi, "uint", 0)
	_WinAPI_ReleaseDC(0, $hDC)
	_WinAPI_DeleteObject($hBmp)
	If @error Or $r[0] <> $h Then
		$gBild = 0
		Return False
	EndIf
	Return True
EndFunc

; Farbe 0xRRGGBB am Bildschirmpunkt $x,$y aus dem eingelesenen Bild, -1 ausserhalb
Func BildFarbe($x, $y)
	Local $i = ($y - $gBildY) * $gBildB + ($x - $gBildX)
	If Not IsDllStruct($gBild) Or $x < $gBildX Or $x >= $gBildX + $gBildB Or $i < 0 Or $i >= DllStructGetSize($gBild) / 4 Then Return -1
	Return BitAND(DllStructGetData($gBild, 1, $i + 1), 0xFFFFFF)
EndFunc

; Bildschirm-X des ersten senkrechten Pfeilschafts der Farbe $farbe auf Hoehe $y zwischen $x1 und $x2,
; mit $arme zusaetzlich gepruefte Pfeilspitze; -1 wenn keiner da ist (liest aus BildLesen)
Func PfeilSchaft($x1, $x2, $y, $farbe, $arme)
	For $x = $x1 To $x2
		If Not FarbeNah($x, $y, $farbe, 10) Then ContinueLoop
		; senkrechter Schaft: von den 4 Nachbarn darueber und darunter darf einer abweichen
		; (der Schaft hat ein etwas dunkleres Pixel)
		If FarbeNah($x, $y - 2, $farbe, 12) + FarbeNah($x, $y - 1, $farbe, 12) _
				+ FarbeNah($x, $y + 1, $farbe, 12) + FarbeNah($x, $y + 2, $farbe, 12) >= 3 Then
			If Not $arme Then Return $x
			; oberes Ende des Schafts suchen, eine Zeile darunter liegen links und rechts die Arme
			Local $o = $y
			While $o > $y - 20 And (FarbeNah($x, $o - 1, $farbe, 12) Or FarbeNah($x, $o - 2, $farbe, 12))
				$o -= 1
			WEnd
			If FarbeNah($x - 2, $o + 1, $KbPfeilArm, 25) And FarbeNah($x + 2, $o + 1, $KbPfeilArm, 25) Then Return $x
		EndIf
	Next
	Return -1
EndFunc

Func FarbeNah($x, $y, $farbe, $tol)
	Local $c = BildFarbe($x, $y)
	If $c < 0 Then Return False
	Return Abs(BitAND(BitShift($c, 16), 255) - BitAND(BitShift($farbe, 16), 255)) <= $tol _
			And Abs(BitAND(BitShift($c, 8), 255) - BitAND(BitShift($farbe, 8), 255)) <= $tol _
			And Abs(BitAND($c, 255) - BitAND($farbe, 255)) <= $tol
EndFunc

; ist die MO-Tagesuebersicht das aktive Fenster? (wird im Tastatur-Hook aufgerufen, muss schnell sein)
Func TagesAktiv()
	Return _WinAPI_GetClassName(_WinAPI_GetForegroundWindow()) = $TagesKlasse
EndFunc

; Tagesuebersicht, offene To-Dos: in der markierten Zeile "Bereich" anklicken, "e" (erledigt) waehlen
; und danach die nachrueckende Zeile markieren. Die Liste (TNewStringGrid) hat keine ansprechbaren
; Zeilen, daher wird die markierte Zeile am schwarzen Rahmen in der farbigen Randspalte erkannt.
; Geklickt wird mit der echten Maus, weil das Menue an der Mausposition aufgeht.
Func Abhaken()
	Local $hWnd = WinGetHandle("[ACTIVE]")
	If _WinAPI_GetClassName($hWnd) <> $TagesKlasse Then Return False
	Local $hGrid = ControlGetHandle($hWnd, "", "[CLASS:TNewStringGrid; INSTANCE:1]")
	If @error Then Return Meldung("Liste der Tagesuebersicht nicht gefunden")
	If Not OffeneToDos($hWnd) Then Return Meldung("Nur bei ""Offene To-Dos"" moeglich")
	; erst loslassen lassen, sonst kaemen Klick und "e" mit gedrueckter Strg+Alt an
	If Not ModifierLos() Then Return Meldung("Strg und Alt bitte loslassen")
	Local $g = WinGetPos($hGrid)
	If @error Then Return False
	Opt("PixelCoordMode", 1)
	Opt("MouseCoordMode", 1)
	Local $mitte = MarkierteZeile($g)
	If $mitte < 0 Then Return Meldung("Keine markierte Zeile sichtbar")
	Local $alt = MouseGetPos()
	MouseClick("left", $g[0] + Int(IniRead($Ini, "Tagesuebersicht", "BereichX", $defBereichX)), $mitte, 1, 0)
	; Kontextmenue (Offen / In Arbeit / erledigt) abwarten
	If Not MenueWarten(True) Then
		MouseMove($alt[0], $alt[1], 0)
		Return Meldung("Menue Offen/In Arbeit/erledigt ist nicht erschienen")
	EndIf
	Send("e")
	MenueWarten(False)
	Sleep(Int(IniRead($Ini, "Tagesuebersicht", "Warten", $defAbhakenWarten)))
	; nachrueckende Zeile markieren; war es die letzte Zeile, ist dort jetzt leere (weisse) Flaeche,
	; dann die daruber
	Local $nameX = $g[0] + Int(IniRead($Ini, "Tagesuebersicht", "NameX", $defNameX))
	If PixelGetColor($nameX, $mitte) = 0xFFFFFF And $mitte - $ZeilenHoehe > $g[1] + $KopfHoehe Then $mitte -= $ZeilenHoehe
	If PixelGetColor($nameX, $mitte) <> 0xFFFFFF Then MouseClick("left", $nameX, $mitte, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	Return True
EndFunc

; wartet bis zu 2 s, bis ein Kontextmenue sichtbar ($offen = True) bzw. keins mehr sichtbar ist;
; unsichtbare Menuefenster haelt Windows oft vorraetig, die zaehlen nicht
Func MenueWarten($offen)
	Local $t = TimerInit()
	While TimerDiff($t) < 2000
		Local $sichtbar = False, $liste = WinList("[CLASS:#32768]")
		For $i = 1 To $liste[0][0]
			If BitAND(WinGetState($liste[$i][1]), 2) Then $sichtbar = True
		Next
		If $sichtbar = $offen Then Return True
		Sleep(20)
	WEnd
	Return False
EndFunc

; Bildschirm-Y der Mitte der markierten Zeile, -1 wenn keine zu sehen ist
Func MarkierteZeile($g)
	Local $x = $g[0] + $RandSpalteX
	Local $p = PixelSearch($x, $g[1] + $KopfHoehe, $x, $g[1] + $g[3] - 3, 0x000000, 16)
	If @error Then Return -1
	; Treffer ist die obere Rahmenlinie, die Zeile ist 20 Pixel hoch
	Return $p[1] + Int($ZeilenHoehe / 2)
EndFunc

; der Knopf "Offene To-Dos" ist eingedrueckt hellblau hinterlegt
Func OffeneToDos($hWnd)
	Local $hKnopf = ControlGetHandle($hWnd, "", "[CLASS:TAdvGlowButton; TEXT:Offene To-Dos]")
	If @error Then Return False
	Local $p = WinGetPos($hKnopf)
	If @error Then Return False
	Opt("PixelCoordMode", 1)
	PixelSearch($p[0] + 3, $p[1] + 3, $p[0] + 5, $p[1] + 5, $OffenFarbe, 30)
	Return Not @error
EndFunc

; wartet bis zu 3 s, bis Strg, Alt, AltGr und Umschalt losgelassen sind
Func ModifierLos()
	Local $t = TimerInit()
	While TimerDiff($t) < 3000
		If Not (Gedrueckt(0x10) Or Gedrueckt(0x11) Or Gedrueckt(0x12)) Then Return True
		Sleep(20)
	WEnd
	Return False
EndFunc

; Strg+Alt+Umsch+A in der Tagesuebersicht: Maus auf die Spalte "Bereich" halten
Func KalibBereich()
	Local $hGrid = ControlGetHandle(WinGetHandle("[ACTIVE]"), "", "[CLASS:TNewStringGrid; INSTANCE:1]")
	If @error Then Return Meldung("Liste der Tagesuebersicht nicht gefunden")
	Opt("MouseCoordMode", 1)
	Local $m = MouseGetPos(), $g = WinGetPos($hGrid)
	Local $x = $m[0] - $g[0]
	If $x < 0 Or $x >= $g[2] Then Return Meldung("Maus steht nicht auf der Liste")
	DirCreate(@AppDataDir & "\MOReiter")
	IniWrite($Ini, "Tagesuebersicht", "BereichX", $x)
	Hinweis("Spalte Bereich kalibriert: " & $x)
	Return True
EndFunc

; Pfeil rechts neben der Patientensuche: der Knopf ganz rechts im Panel des Suchfelds
; (die Knoepfe sind TAdvSmoothButton ohne Text, daher Auswahl ueber die Lage)
Func LetztePatienten()
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $hSuche = ControlGetHandle($hWnd, "", "[CLASS:TAlnumEdit; INSTANCE:1]")
	If @error Then Return Meldung("Patientensuche nicht gefunden")
	Local $hKnopf = RandKnopf(_WinAPI_GetParent($hSuche), "rechts")
	If Not $hKnopf Then Return Meldung("Pfeil neben der Patientensuche nicht gefunden")
	Return KnopfKlick($hWnd, $hKnopf)
EndFunc

; drei Striche links oben: der oberste, linkeste Knopf im Kopfbereich
Func Menue()
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $hKopf = ControlGetHandle($hWnd, "", "[CLASS:THeaderForm; INSTANCE:1]")
	If @error Then Return Meldung("Kopfbereich nicht gefunden")
	Local $hKnopf = RandKnopf($hKopf, "linksoben")
	If Not $hKnopf Then Return Meldung("Menueknopf nicht gefunden")
	Return KnopfKlick($hWnd, $hKnopf)
EndFunc

; sichtbarer TAdvSmoothButton unterhalb von $hEltern, der am weitesten rechts bzw. links oben liegt
Func RandKnopf($hEltern, $wo)
	Local $liste = _WinAPI_EnumChildWindows($hEltern)
	If @error Then Return 0
	Local $best = 0, $bestWert = 0
	For $i = 1 To $liste[0][0]
		If $liste[$i][1] <> "TAdvSmoothButton" Then ContinueLoop
		Local $p = WinGetPos($liste[$i][0])
		If @error Then ContinueLoop
		Local $wert = ($wo = "rechts") ? $p[0] : -($p[1] * 100000 + $p[0])
		If $best = 0 Or $wert > $bestWert Then
			$best = $liste[$i][0]
			$bestWert = $wert
		EndIf
	Next
	Return $best
EndFunc

; klickt in die Mitte des Knopfs $hKnopf
Func KnopfKlick($hWnd, $hKnopf)
	Local $p = WinGetPos($hKnopf)
	If @error Then Return Meldung("Knopf nicht gefunden")
	WinActivate($hWnd)
	WinWaitActive($hWnd, "", 2)
	If IniRead($Ini, "Allgemein", "EchteMaus", "0") = "1" Then
		Local $alt = MouseGetPos()
		Opt("MouseCoordMode", 1)
		MouseClick("left", $p[0] + Int($p[2] / 2), $p[1] + Int($p[3] / 2), 1, 0)
		MouseMove($alt[0], $alt[1], 0)
		Return True
	EndIf
	If Not ControlClick($hWnd, "", $hKnopf, "left", 1, Int($p[2] / 2), Int($p[3] / 2)) Then Return Meldung("Klick fehlgeschlagen")
	Return True
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
