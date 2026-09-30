#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=MOReiter.ico
#AutoIt3Wrapper_Outfile=MOReiter.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Kompilieren mit Icon: Aut2exe.exe /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico
; MOReiter: waehlt in Medical Office den Reiter "Kartei", "Krankenblatt", "ePA" oder "ePAAbr" per Mausklick
; Aufruf:  MOReiter.exe Kartei | Krankenblatt | ePA | ePAAbr  -> einmal klicken und beenden (z.B. aus VB6 per Shell)
;          MOReiter.exe LetztePatienten | Menue  -> Pfeil neben der Patientensuche bzw. drei Striche links oben
;          MOReiter.exe Filter n                 -> Kartei und dort den n-ten Filter waehlen
;          MOReiter.exe Diagnose n               -> in der Diagnoseerfassung die n-te Kurzwahl rechts anklicken
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
;            Alt+Pfeil hoch/runter  nur in der MO-Tagesuebersicht: in der Bereichsauswahl links den Eintrag
;                              ueber bzw. unter dem gelb umrandeten anklicken (vom obersten zum untersten
;                              und umgekehrt, der Trennstrich wird uebersprungen), danach rechts die oberste
;                              Zeile der Liste (geklickt wird nach dem Loslassen von Alt)
;            Strg+Alt+Umsch+K  Kalibrieren: Maus auf Reiter "Kartei" halten und druecken
;            Strg+Alt+Umsch+L  Kalibrieren: Maus auf Reiter "Krankenblatt" halten und druecken
;            Strg+Alt+Umsch+P  Kalibrieren: Maus auf Reiter "ePA" halten und druecken
;            Strg+Alt+Umsch+A  Kalibrieren: Maus auf Reiter "ePAAbr" halten und druecken
;            Strg+Alt+ ^ 1 2 3 4 5 6 7 8 9 0 sz Akut F2 .. F12
;                              Kartei (Reiter in [Filter] Reiter=) und dort der 1., 2., ... 24. Filter
;                              links ("Alle Eintraege", "Dateien", "RR Gewicht", ...)
;            Strg+Alt+Umsch+ dieselbe Taste  Kalibrieren: Maus auf diesen Filter halten und druecken;
;                              beim 1. Filter (^) wird die Position gemerkt, bei den anderen der Zeilenabstand
;            Strg+Alt+ ^ 1 2 ... F12 in der Diagnoseerfassung: stattdessen die 1., 2., ... 24. Diagnose-
;                              Kurzwahl rechts (erst die linke Spalte von oben nach unten, dann die naechste;
;                              [Diagnosen] Reihenfolge=Zeilen zaehlt zeilenweise); ohne Kalibrieren, die
;                              Knoepfe werden bei jedem Druck am Bildschirm gesucht
;            Strg+Alt halten und auf dem Ziffernblock eine Nummer tippen (bis 3 Stellen, mit und ohne
;                              NumLock): beim Loslassen von Strg bzw. Alt wird der Filter bzw. in der
;                              Diagnoseerfassung die Kurzwahl mit dieser Nummer gewaehlt, auch jenseits von 24;
;                              gezaehlt ab 0 wie die Tasten ^ 1 2 ..., so dass Ziffernblock 1-9 dasselbe waehlt
;                              wie die Tasten 1-9 (0 = ^, 10 = Taste 0, 11 = sz, 12 = Akut, 13 = F2 ...)
;            ohne Ziffernblock: Strg+Alt halten und mehrere Ziffern der oberen Reihe tippen, z.B. 1 7:
;                              in der Kartei wird sofort Filter 1 gewaehlt und mit der 7 dann Filter 17
;                              (wie Ziffernblock 17); in der Diagnoseerfassung wird erst beim Loslassen
;                              gewaehlt, damit nicht nebenbei Kurzwahl 1 angeklickt wird (eine einzelne 0
;                              bleibt dort die Taste 0, also Nr. 10). Die Folge endet mit dem Loslassen
;                              von Strg oder Alt, nach 3 Ziffern oder nach [Allgemein] FolgeZeit ms Pause
;            Strg+Alt eine halbe Sekunde halten ([Allgemein] Einblendung=ms, 0 = nie): an den Filtern
;                              der Kartei bzw. den Diagnose-Kurzwahlen erscheinen gelbe Schildchen mit
;                              Ziffernblock-Nummer, dahinter die direkte Taste, wenn sie anders heisst
;                              (z.B. "0 ^", "5", "11 sz"), im Briefversand mit dem Buchstaben der Taste an
;                              der Klickstelle, bis Strg oder Alt losgelassen wird
;            Die Tasten werden per Tastatur-Hook abgefangen (nicht per HotKeySet), weil nur so linkes Alt
;            von AltGr unterschieden werden kann: AltGr kommt als linke Strg + rechte Alt an und bleibt
;            unberuehrt, so dass AltGr+2 3 7 8 9 0 sz weiterhin hoch2 hoch3 { [ ] } \ liefern.
;            (Das fruehere Durchreichen per Send liess gelegentlich die Strg-Taste haengen.)
;            im Fenster "MEDICAL OFFICE - Briefversand" statt der obigen Belegung:
;            Strg+Alt+K Kontaktverzeichnis, I KIM-Verzeichnis, R Krankenblatt, O Ordner (Anhaenge),
;                              E Kaestchen "Empfangsbestaetigung anfordern", B Betreff-Zeile, M oberster
;                              Empfaenger, F Brief-Vorschau, V Versenden/Drucken; Esc allein: Abbrechen
;                              (geklickt wird mit der echten Maus, nach dem Loslassen von Strg und Alt);
;                              beim Halten von Strg+Alt erscheinen die Buchstaben gelb an den Klickstellen
;            Strg+Alt+Umsch+D  Liste der Controls im aktiven Fenster (Klasse, Nummer, Text, Lage) in die
;                              Zwischenablage und nach %APPDATA%\MOReiter\Fenster.txt, zum Einrichten neuer Tasten
;            Beenden ueber das Tray-Menue (kein Strg+Alt+Q, das waere AltGr+Q = @)
; Laeuft Medical Office noch nicht, startet jeder Aufruf die Zentrale (medoff.exe) zum Anmelden.
; Die Reiterleiste ist ein Delphi-Control (TmoTabSet) ohne eigene Handles je Reiter,
; deshalb wird relativ zur linken oberen Ecke dieses Controls geklickt.
#include <Misc.au3>
#include <WinAPI.au3>
#include <WindowsNotifsConstants.au3>
#include <ScreenCapture.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
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
; Bereichsauswahl links (gemessen 30.9.26): der gewaehlte Eintrag hat 2 Pixel eingerueckt einen gelben
; Rahmen (8F8F1E, innen E2E178); die blaue Hervorhebung unter der Maus kann auf jedem Eintrag liegen und
; auch den gewaehlten bis auf einen gelben Rand links und rechts ueberdecken, deshalb zaehlt nur der Rahmen.
; Die Eintraege haben einen grauen Rand (C1C6CF) und 1 Pixel Abstand; der Trennstrich ist innen weiss.
; Briefversand: Titel des Fensters; Betreff, Empfaengerliste und Vorschau werden nach der Klasse
; vermutet, falls in [Briefversand] kein Control angegeben ist; erste Empfaengerzeile relativ zur Liste
Const $BriefTitel = "MEDICAL OFFICE - Briefversand"
Const $defEmpfaengerX = 150, $defEmpfaengerY = 33
Const $defBereichRahmen = 0x8F8F1E, $BereichRand = 0xC1C6CF, $defBereichWarten = 1500
; Krankenblatt-Liste (gemessen 28.9.26): Hintergrund der markierten Zeile, Kopfzeilenhoehe, Hoehe des
; Pfeilschafts ab Zeilenoberkante; Pfeil nach oben hellblau = noch nicht in der ePA, dunkel = schon drin
Const $KbMarkiertFarbe = 0xD3E3F7, $KbKopfHoehe = 18, $KbPfeilY = 11
Const $KbPfeilHell = 0x6CC1EF, $KbPfeilDunkel = 0x233658, $KbPfeilArm = 0x809FBA
; Fenster nach dem Pfeilklick, darin wird der Knopf "Hochladen" nach $defHochladenWarten ms gedrueckt
; (100 ms reichten nur, solange nach jedem Fensterbefehl noch 250 ms WinWaitDelay dazukamen)
Const $EpaDialog = "[REGEXPTITLE:^MEDICAL OFFICE - ePA; CLASS:TClientWindowForm]", $defHochladenWarten = 300
; danach evtl. Dialog zur Leistungsdokumentation mit "Uebernehmen", wird bis zu $defLeistungWarten ms erwartet
Const $LeistungDialog = "[TITLE:Medical Office; CLASS:TfrmDialogContainer]", $defLeistungWarten = 1500
; Diagnoseerfassung (gemessen 30.9.26): Kurzwahl-Knoepfe rechts grau (182) auf etwas hellerem Grau (192),
; Toleranz je Farbanteil; der Knopf mit dem Fokus ist dunkler und blau umrandet
Const $DiagFenster = "[REGEXPTITLE:^Diagnoseerfassung]", $DiagTitel = "Diagnoseerfassung"
Const $defDiagKnopfFarbe = 0xB6B6B6, $defDiagToleranz = 6, $defDiagReihenfolge = "Spalten"
; virtuelle Tastencodes fuer die Filter 1..24 (deutsches Layout: ^ = OEM_5, sz = OEM_4, Akut = OEM_6)
Global $FilterTasten[24] = [0xDC, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30, 0xDB, 0xDD, _
		0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7A, 0x7B]
; Beschriftung dieser Tasten fuer die Einblendung
Global $TastenNamen[24] = ["^", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ChrW(223), ChrW(180), _
		"F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]
Const $defEinblendung = 500 ; ms Strg+Alt halten, bis die Nummern eingeblendet werden
; Ziffernblock: im Hook gesammelte Nummer, ihr Ziel ("Filter"/"Diagnose"), geschluckte, noch
; gedrueckte Ziffertasten als ",vk,", und die gerade angezeigte Nummer
Global $gNummer = "", $gNummerZiel = "", $gUnten = "", $gNummerAngezeigt = ""
; die Nummer stammt von der oberen Ziffernreihe (dann heisst eine einzelne 0 Nr. 10, wie die Taste 0)
Global $gNummerOben = False
; Kartei: bei gehaltenem Strg+Alt bisher getippte Ziffern der oberen Reihe und Zeit der letzten
Global $gFolge = "", $gFolgeZeit = 0
Const $defFolgeZeit = 2000 ; ms, eine laengere Pause zwischen zwei Ziffern beginnt eine neue Folge
; Einblendung: Fenster mit den Schildchen, zugehoeriges Fenster und dessen Kurzwahl-Knoepfe,
; Beginn des Haltens von Strg+Alt (0 = nicht gehalten), schon versucht
Global $gEinblendung = 0, $gEinblendungWnd = 0, $gEinblendungKnoepfe = 0, $gHaltStart = 0, $gEinblendungVersucht = False
; Tastatur-Hook: der Rueckruf merkt sich nur die Aufgabe, ausgefuehrt wird sie in der Hauptschleife,
; weil Windows einen Hook, der zu lange braucht, stillschweigend abhaengt
Global $gHook = 0, $gAufgabe = "", $gGehalten = 0, $gGehaltenZeit = 0
; eingelesener Bildschirmausschnitt fuer die Pfeilpruefung (BildLesen/BildFarbe)
Global $gBild = 0, $gBildX = 0, $gBildY = 0, $gBildB = 0

If $CmdLine[0] > 1 And $CmdLine[1] = "Filter" Then
	Exit FilterWahl(Int($CmdLine[2]) - 1) ? 0 : 1
ElseIf $CmdLine[0] > 1 And $CmdLine[1] = "Diagnose" Then
	Exit DiagnoseWahl(Int($CmdLine[2]) - 1) ? 0 : 1
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
TraySetToolTip("MOReiter: Strg+Alt+K Kartei/Wechsel, L Krankenblatt, P ePA/ePAAbr, Z letzte Patienten, Leertaste Menue, ^..F12 oder Ziffernblock Filter bzw. Diagnose-Kurzwahl")
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
		$gHaltStart = 0
		$gEinblendungVersucht = False
	EndIf
	NummerPruefen()
	EinblendungPruefen()
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
			If StringInStr($gUnten, "," & $vk & ",") Then
				$gUnten = StringReplace($gUnten, "," & $vk & ",", "", 1)
				Return 1
			EndIf
			; Strg bzw. linke Alt losgelassen: Ziffernfolge der oberen Reihe ist zu Ende
			If $vk = 0xA2 Or $vk = 0xA3 Or $vk = 0xA4 Then $gFolge = ""
			If $vk = $gGehalten Then
				$gGehalten = 0
				Return 1
			EndIf
		Else
			; Ziffernblock bei Strg+Alt: Ziffer an die Nummer haengen, gewaehlt wird beim Loslassen
			; (NummerPruefen); geschluckt, damit auch keine Alt+Ziffernblock-Zeichencodes entstehen
			; in der Diagnoseerfassung ebenso die obere Ziffernreihe (ohne Umschalt, das ist dort frei)
			Local $z = ZifferVon($vk, $kb.flags), $oben = False
			If $z < 0 And $vk >= 0x30 And $vk <= 0x39 And StrgAlt() And DiagAktiv() And Not Gedrueckt(0x10) Then
				$z = $vk - 0x30
				$oben = True
			EndIf
			If $z >= 0 And StrgAlt() Then
				; schon gedrueckt: automatische Wiederholung
				If Not StringInStr($gUnten, "," & $vk & ",") Then
					$gUnten &= "," & $vk & ","
					If $gNummer = "" Then
						$gNummerZiel = DiagAktiv() ? "Diagnose" : "Filter"
						$gNummerOben = $oben
					EndIf
					If StringLen($gNummer) < 3 Then $gNummer &= $z
				EndIf
				Return 1
			EndIf
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
				$gAufgabe = FolgeAufgabe($vk, $aufgabe)
				Return 1
			EndIf
		EndIf
	EndIf
	Return _WinAPI_CallNextHookEx($gHook, $nCode, $wParam, $lParam)
EndFunc

; Aufgabe zur Taste $vk, falls gerade Strg + linke Alt (ohne AltGr) gehalten werden, sonst ""
Func HookAufgabe($vk)
	; Alt+Pfeil hoch/runter (ohne Strg und Umschalt) nur in der Tagesuebersicht
	If ($vk = 0x26 Or $vk = 0x28) And Gedrueckt(0x12) And Not Gedrueckt(0x11) And Not Gedrueckt(0x10) _
			And TagesAktiv() Then Return ($vk = 0x26) ? "BereichHoch" : "BereichRunter"
	; Esc allein im Briefversand: Abbrechen
	If $vk = 0x1B And Not (Gedrueckt(0x10) Or Gedrueckt(0x11) Or Gedrueckt(0x12) Or Gedrueckt(0x5B) Or Gedrueckt(0x5C)) _
			And BriefAktiv() Then Return "Brief:Abbrechen"
	If Not StrgAlt() Then Return ""
	Local $kalib = Gedrueckt(0x10)
	If $vk = 0x44 And $kalib Then Return "Fensterliste" ; D
	If BriefAktiv() Then
		Local $brief = BriefTaste($vk)
		If $brief <> "" Then Return $kalib ? "" : "Brief:" & $brief
	EndIf
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
		If $vk <> $FilterTasten[$i] Then ContinueLoop
		; in der Diagnoseerfassung waehlen dieselben Tasten die Diagnose-Kurzwahl (nichts zu kalibrieren)
		If DiagAktiv() Then Return $kalib ? "" : "Diagnose:" & $i
		Return ($kalib ? "KalibFilter:" : "Filter:") & $i
	Next
	Return ""
EndFunc

; Strg + linke Alt gehalten, aber nicht AltGr (kommt als linke Strg + rechte Alt)
Func StrgAlt()
	Return (Gedrueckt(0xA2) Or Gedrueckt(0xA3)) And Gedrueckt(0xA4) And Not Gedrueckt(0xA5)
EndFunc

; Ziffer 0-9 einer Ziffernblocktaste, -1 fuer andere Tasten; ohne NumLock liefert der Ziffernblock
; Pfeil- und Blaettertasten, die sich vom eigenen Pfeilblock nur durch das fehlende Extended-Bit
; (LLKHF_EXTENDED = 1) unterscheiden
Func ZifferVon($vk, $flags)
	If $vk >= 0x60 And $vk <= 0x69 Then Return $vk - 0x60
	If BitAND($flags, 1) Then Return -1
	Switch $vk
		Case 0x2D ; Einfg
			Return 0
		Case 0x23 ; Ende
			Return 1
		Case 0x28 ; Pfeil runter
			Return 2
		Case 0x22 ; Bild runter
			Return 3
		Case 0x25 ; Pfeil links
			Return 4
		Case 0x0C ; Clear
			Return 5
		Case 0x27 ; Pfeil rechts
			Return 6
		Case 0x24 ; Pos1
			Return 7
		Case 0x26 ; Pfeil hoch
			Return 8
		Case 0x21 ; Bild hoch
			Return 9
	EndSwitch
	Return -1
EndFunc

; Kartei, obere Ziffernreihe: die erste Ziffer waehlt wie gewohnt sofort (Taste 1 = Filter 1), jede
; weitere bei durchgehend gehaltenem Strg+Alt haengt sich an und waehlt den Filter mit der ganzen
; Nummer (1, 7 = Filter 17 wie Ziffernblock 17); liefert die auszufuehrende Aufgabe
Func FolgeAufgabe($vk, $aufgabe)
	If $vk < 0x30 Or $vk > 0x39 Or StringLeft($aufgabe, 7) <> "Filter:" Then
		$gFolge = ""
		Return $aufgabe
	EndIf
	Local $d = String($vk - 0x30)
	If $gFolge <> "" And StringLen($gFolge) < 3 _
			And TimerDiff($gFolgeZeit) < Int(IniRead($Ini, "Allgemein", "FolgeZeit", $defFolgeZeit)) Then
		$gFolge &= $d
		$aufgabe = "Filter:" & Int($gFolge)
	Else
		$gFolge = $d
	EndIf
	$gFolgeZeit = TimerInit()
	Return $aufgabe
EndFunc

; zeigt die getippte Nummer an und waehlt sie, sobald Strg oder Alt losgelassen ist
Func NummerPruefen()
	If $gNummer = "" Then Return
	If StrgAlt() Then
		If $gNummer <> $gNummerAngezeigt Then
			$gNummerAngezeigt = $gNummer
			AdlibUnRegister("HinweisWeg")
			ToolTip("Nr. " & $gNummer, Default, Default, "MOReiter")
		EndIf
		Return
	EndIf
	Local $n = Int($gNummer), $ziel = $gNummerZiel
	; einzelne 0 der oberen Reihe: wie die Taste 0 die 11. Kurzwahl
	If $gNummerOben And $gNummer = "0" Then $n = 10
	$gNummer = ""
	$gNummerAngezeigt = ""
	$gUnten = ""
	ToolTip("")
	; ab 0 gezaehlt wie die Tasten ^ 1 2 ...
	If $ziel = "Diagnose" Then Return DiagnoseWahl($n)
	Return FilterWahl($n)
EndFunc

; blendet nach $defEinblendung ms Halten von Strg+Alt (ohne Umschalt) die Nummern ein, beim
; Loslassen wieder aus; je Halten nur ein Versuch, weil die Knopfsuche Zeit kostet
Func EinblendungPruefen()
	If Not StrgAlt() Or Gedrueckt(0x10) Then
		$gHaltStart = 0
		$gEinblendungVersucht = False
		If $gEinblendung Then EinblendungWeg()
		Return
	EndIf
	If $gHaltStart = 0 Then $gHaltStart = TimerInit()
	Local $ms = Int(IniRead($Ini, "Allgemein", "Einblendung", $defEinblendung))
	If $ms <= 0 Or $gEinblendungVersucht Or TimerDiff($gHaltStart) < $ms Then Return
	$gEinblendungVersucht = True
	EinblendungZeigen()
EndFunc

; Schildchen "Nummer Taste" links an jeder Diagnose-Kurzwahl bzw. an jedem Filter der Kartei; ein
; einziges durchsichtiges, nicht anklickbares Fenster (Farbschluessel Magenta) ueber dem Zielfenster
Func EinblendungZeigen()
	; $pos: linker Rand des Schildchens, Mitte y, Text ("" = Nummer und Taste aus der Stelle)
	Local $hWnd = _WinAPI_GetForegroundWindow(), $pos[0][3], $n = 0
	If BriefAktiv() Then
		$n = BriefPositionen($hWnd, $pos)
	ElseIf DiagAktiv() Then
		Local $k = DiagKnoepfe($hWnd)
		If Not IsArray($k) Then Return
		$gEinblendungKnoepfe = $k
		ReDim $pos[UBound($k)][3]
		For $i = 0 To UBound($k) - 1
			$pos[$i][0] = $k[$i][3] + 2
			$pos[$i][1] = $k[$i][1]
		Next
		$n = UBound($k)
	ElseIf MOAktiv() Then
		$n = FilterPositionen($hWnd, $pos)
	EndIf
	If $n = 0 Then Return
	Local $f = WinGetPos($hWnd)
	If @error Then Return
	Local $g = GUICreate("MOReiter-Einblendung", $f[2], $f[3], $f[0], $f[1], $WS_POPUP, _
			BitOR($WS_EX_LAYERED, $WS_EX_TRANSPARENT, $WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_NOACTIVATE))
	GUISetBkColor(0xFF00FF, $g)
	; 1 = LWA_COLORKEY: Magenta ist durchsichtig
	DllCall("user32.dll", "bool", "SetLayeredWindowAttributes", "hwnd", $g, "dword", 0xFF00FF, "byte", 255, "dword", 1)
	For $i = 0 To $n - 1
		; Ziffernblock-Nummer (ab 0), dahinter die direkte Taste, falls sie nicht gleich heisst
		Local $t = $pos[$i][2]
		If $t = "" Then
			$t = String($i)
			If $i < UBound($TastenNamen) And $TastenNamen[$i] <> $t Then $t &= " " & $TastenNamen[$i]
		EndIf
		GUICtrlCreateLabel($t, $pos[$i][0] - $f[0], $pos[$i][1] - $f[1] - 8, 7 * StringLen($t) + 8, 16, BitOR($SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, 0xFFE45C)
		GUICtrlSetColor(-1, 0x000000)
		GUICtrlSetFont(-1, 8.5, 700, 0, "Segoe UI")
	Next
	GUISetState(@SW_SHOWNOACTIVATE, $g)
	$gEinblendung = $g
	$gEinblendungWnd = $hWnd
EndFunc

Func EinblendungWeg()
	If $gEinblendung Then GUIDelete($gEinblendung)
	$gEinblendung = 0
	$gEinblendungWnd = 0
	$gEinblendungKnoepfe = 0
EndFunc

; Bildschirmpunkte (linker Rand der Filterliste, Mitte des Filters) aller sichtbaren Filter in $pos,
; liefert ihre Anzahl; nur wenn die Kartei (bzw. [Filter] Reiter) schon aktiv ist
Func FilterPositionen($hWnd, ByRef $pos)
	Local $hCtrl = ControlGetHandle($hWnd, "", IniRead($Ini, "Allgemein", "Control", $defCtrl))
	If @error Or Not BitAND(WinGetState($hCtrl), 2) Then Return 0
	If Not ReiterAktiv(IniRead($Ini, "Filter", "Reiter", $defFilterReiter), $hWnd, $hCtrl) Then Return 0
	Local $p = WinGetPos($hCtrl)
	If @error Then Return 0
	Local $x = $p[0] + Int(IniRead($Ini, "Filter", "X", $defFilterX))
	Local $y0 = $p[1] + Int(IniRead($Ini, "Filter", "Y", $defFilterY))
	Local $abst = Number(IniRead($Ini, "Filter", "Abstand", $defFilterAbstand))
	If $abst < 5 Then Return 0
	Local $pt = DllStructCreate("int X;int Y"), $hListe = 0, $links = 0, $unten = 0, $n = 0
	; so lange, wie der Punkt noch auf derselben Filterliste liegt
	While $n < 200
		Local $y = Round($y0 + $n * $abst)
		$pt.X = $x
		$pt.Y = $y
		Local $h = _WinAPI_WindowFromPoint($pt)
		If $n = 0 Then
			If $h = 0 Or _WinAPI_GetAncestor($h, 2) <> $hWnd Then Return 0
			$hListe = $h
			Local $r = WinGetPos($h)
			If @error Then Return 0
			$links = $r[0] + 2
			; bis zum unteren Rand der Liste, abzueglich einer halben Zeile
			$unten = $r[1] + $r[3] - $abst / 2
		ElseIf $h <> $hListe Or $y > $unten Then
			ExitLoop
		EndIf
		ReDim $pos[$n + 1][3]
		$pos[$n][0] = $links
		$pos[$n][1] = $y
		$n += 1
	WEnd
	Return $n
EndFunc

Func Gedrueckt($vk)
	Local $r = DllCall("user32.dll", "short", "GetAsyncKeyState", "int", $vk)
	Return Not @error And BitAND($r[0], 0x8000) <> 0
EndFunc

Func Ausfuehren($aufgabe)
	; die Schildchen stoeren sonst Bildschirmpruefungen; Diagnose/Filter nehmen sie selbst weg
	If StringLeft($aufgabe, 9) <> "Diagnose:" And StringLeft($aufgabe, 7) <> "Filter:" Then EinblendungWeg()
	If $aufgabe = "LetztePatienten" Then Return LetztePatienten()
	If $aufgabe = "Menue" Then Return Menue()
	If $aufgabe = "Abhaken" Then Return Abhaken()
	If $aufgabe = "BereichHoch" Then Return BereichWechsel(-1)
	If $aufgabe = "BereichRunter" Then Return BereichWechsel(1)
	If $aufgabe = "Hochladen" Then Return Hochladen()
	If $aufgabe = "Fensterliste" Then Return FensterListe()
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
		Case "Diagnose"
			DiagnoseWahl(Int($teil[1]))
		Case "Brief"
			BriefKlick($teil[1])
	EndSwitch
EndFunc

; ist der Briefversand das aktive Fenster? (im Hook, muss schnell sein)
Func BriefAktiv()
	Return StringLeft(_WinAPI_GetWindowText(_WinAPI_GetForegroundWindow()), StringLen($BriefTitel)) = $BriefTitel
EndFunc

; Aufgabe zur Taste im Briefversand, "" wenn die Taste dort nicht belegt ist
Func BriefTaste($vk)
	Switch $vk
		Case 0x4B ; K
			Return "Kontakt"
		Case 0x49 ; I
			Return "KIM"
		Case 0x52 ; R
			Return "Krankenblatt"
		Case 0x4F ; O
			Return "Ordner"
		Case 0x45 ; E
			Return "Empfang"
		Case 0x42 ; B
			Return "Betreff"
		Case 0x4D ; M
			Return "Empfaenger"
		Case 0x46 ; F
			Return "Vorschau"
		Case 0x56 ; V
			Return "Versenden"
	EndSwitch
	Return ""
EndFunc

; Briefversand: klickt mit der echten Maus auf das zu $was gehoerende Control, erst nach dem
; Loslassen von Strg und Alt (sonst kaeme ein Strg+Alt+Klick an)
Func BriefKlick($was)
	Local $hWnd = _WinAPI_GetForegroundWindow()
	If StringLeft(_WinAPI_GetWindowText($hWnd), StringLen($BriefTitel)) <> $BriefTitel Then Return False
	Local $x, $y
	If Not BriefZiel($hWnd, $was, $x, $y) Then Return Meldung($was & " im Briefversand nicht gefunden (Strg+Alt+Umsch+D listet die Controls)")
	If Not ModifierLos() Then Return Meldung("Strg und Alt bitte loslassen")
	Opt("MouseCoordMode", 1)
	Local $alt = MouseGetPos()
	MouseClick("left", $x, $y, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	Return True
EndFunc

; Schildchen fuer die Einblendung im Briefversand: Buchstabe der Taste an der Klickstelle
Func BriefPositionen($hWnd, ByRef $pos)
	Local $tasten[10][2] = [["Kontakt", "K"], ["KIM", "I"], ["Krankenblatt", "R"], ["Ordner", "O"], _
			["Empfang", "E"], ["Betreff", "B"], ["Empfaenger", "M"], ["Vorschau", "F"], ["Versenden", "V"], _
			["Abbrechen", "Esc"]]
	Local $n = 0, $x, $y
	For $i = 0 To UBound($tasten) - 1
		If Not BriefZiel($hWnd, $tasten[$i][0], $x, $y) Then ContinueLoop
		ReDim $pos[$n + 1][3]
		; Schildchen mittig ueber der Klickstelle
		$pos[$n][0] = $x - Int((7 * StringLen($tasten[$i][1]) + 8) / 2)
		$pos[$n][1] = $y
		$pos[$n][2] = $tasten[$i][1]
		$n += 1
	Next
	Return $n
EndFunc

; Bildschirmpunkt $x,$y, auf den fuer $was im Briefversand geklickt wird; False, wenn nicht gefunden
Func BriefZiel($hWnd, $was, ByRef $x, ByRef $y)
	Local $h = 0, $p
	$x = -1
	$y = -1
	Switch $was
		Case "Kontakt"
			$h = TextControl($hWnd, "Kontaktverzeichnis")
		Case "KIM"
			$h = TextControl($hWnd, "KIM-Verzeichnis")
		Case "Krankenblatt"
			$h = TextControl($hWnd, "Krankenblatt")
		Case "Ordner"
			$h = TextControl($hWnd, "Ordner")
		Case "Versenden"
			$h = TextControl($hWnd, "Versenden/Drucken")
		Case "Abbrechen"
			$h = TextControl($hWnd, "Abbrechen")
		Case "Empfang"
			; das Kaestchen sitzt am linken Rand des Controls
			$h = TextControl($hWnd, "Empfangsbest" & ChrW(228) & "tigung anfordern")
			If $h Then
				$p = WinGetPos($h)
				$x = $p[0] + 7
				$y = $p[1] + Int($p[3] / 2)
			EndIf
		Case "Betreff", "Empfaenger", "Vorschau"
			$h = BriefControl($hWnd, $was)
			If $h And $was = "Empfaenger" Then
				; in die erste Zeile, rechts vom Kaestchen, damit es nicht umgeschaltet wird
				$p = WinGetPos($h)
				$x = $p[0] + Int(IniRead($Ini, "Briefversand", "EmpfaengerX", $defEmpfaengerX))
				$y = $p[1] + Int(IniRead($Ini, "Briefversand", "EmpfaengerY", $defEmpfaengerY))
			EndIf
	EndSwitch
	If Not $h Then Return False
	If $x < 0 Then
		$p = WinGetPos($h)
		If @error Then Return False
		$x = $p[0] + Int($p[2] / 2)
		$y = $p[1] + Int($p[3] / 2)
	EndIf
	Return True
EndFunc

; sichtbares Control mit genau diesem Text, 0 wenn keins
Func TextControl($hWnd, $text)
	Local $h = ControlGetHandle($hWnd, "", "[TEXT:" & $text & "]")
	If @error Or Not BitAND(WinGetState($h), 2) Then Return 0
	Return $h
EndFunc

; Betreff-Zeile, Empfaengerliste bzw. Vorschau: aus [Briefversand] $was=<Control in AutoIt-Schreibweise>,
; sonst vermutet: Betreff das oberste einzeilige Eingabefeld, Empfaenger die oberste Liste,
; Vorschau das groesste Control ohne Unterfenster unterhalb des Betreffs in der linken Haelfte
Func BriefControl($hWnd, $was)
	Local $ctl = IniRead($Ini, "Briefversand", $was, "")
	If $ctl <> "" Then
		Local $hc = ControlGetHandle($hWnd, "", $ctl)
		If @error Or Not BitAND(WinGetState($hc), 2) Then Return 0
		Return $hc
	EndIf
	Local $liste = _WinAPI_EnumChildWindows($hWnd)
	If @error Then Return 0
	Local $f = WinGetPos($hWnd), $best = 0, $bestWert = 0, $betreffY = -1
	If $was = "Vorschau" Then
		Local $hb = BriefControl($hWnd, "Betreff")
		If $hb Then
			Local $pb = WinGetPos($hb)
			$betreffY = $pb[1] + $pb[3]
		EndIf
	EndIf
	For $i = 1 To $liste[0][0]
		Local $h = $liste[$i][0], $kl = $liste[$i][1]
		If Not BitAND(WinGetState($h), 2) Then ContinueLoop
		Local $p = WinGetPos($h)
		If @error Then ContinueLoop
		Local $wert = 0
		Switch $was
			Case "Betreff"
				If StringRegExp($kl, "(?i)edit") And Not StringRegExp($kl, "(?i)memo|rich|inner") _
						And $p[2] >= 150 And $p[3] <= 40 Then $wert = 100000 - $p[1]
			Case "Empfaenger"
				If StringRegExp($kl, "(?i)grid|list|tree") And $p[3] >= 40 Then $wert = 100000 - $p[1]
			Case "Vorschau"
				; 5 = GW_CHILD: nur Controls ohne Unterfenster
				If _WinAPI_GetWindow($h, 5) = 0 And $p[1] > $betreffY _
						And $p[0] + $p[2] / 2 < $f[0] + $f[2] / 2 Then $wert = $p[2] * $p[3]
		EndSwitch
		If $wert > $bestWert Then
			$best = $h
			$bestWert = $wert
		EndIf
	Next
	Return $best
EndFunc

; Strg+Alt+Umsch+D: alle Controls des aktiven Fensters mit Klasse, AutoIt-Nummer (INSTANCE), sichtbar,
; Lage relativ zum Fenster und Text in die Zwischenablage und nach Fenster.txt
Func FensterListe()
	Local $hWnd = _WinAPI_GetForegroundWindow()
	Local $f = WinGetPos($hWnd)
	If @error Then Return False
	Local $txt = "Fenster: """ & _WinAPI_GetWindowText($hWnd) & """  Klasse " & _WinAPI_GetClassName($hWnd) _
			& "  Lage " & $f[0] & "," & $f[1] & " " & $f[2] & "x" & $f[3] & @CRLF
	Local $liste = _WinAPI_EnumChildWindows($hWnd, False)
	If Not @error Then
		; INSTANCE zaehlt je Klasse in der Reihenfolge dieser Aufzaehlung
		Local $nr = ObjCreate("Scripting.Dictionary")
		For $i = 1 To $liste[0][0]
			Local $h = $liste[$i][0], $kl = $liste[$i][1]
			$nr.Item($kl) = $nr.Item($kl) + 1
			Local $p = WinGetPos($h)
			If @error Then ContinueLoop
			$txt &= "[CLASS:" & $kl & "; INSTANCE:" & $nr.Item($kl) & "]" _
					& (BitAND(WinGetState($h), 2) ? "" : " (unsichtbar)") _
					& "  " & ($p[0] - $f[0]) & "," & ($p[1] - $f[1]) & " " & $p[2] & "x" & $p[3] _
					& "  """ & StringLeft(_WinAPI_GetWindowText($h), 60) & """" & @CRLF
		Next
	EndIf
	DirCreate(@AppDataDir & "\MOReiter")
	Local $datei = @AppDataDir & "\MOReiter\Fenster.txt"
	; 2 = ueberschreiben, 128 = UTF-8 mit BOM
	Local $fh = FileOpen($datei, 2 + 128)
	FileWrite($fh, $txt)
	FileClose($fh)
	ClipPut($txt)
	Hinweis("Controls in der Zwischenablage und in " & $datei)
	Return True
EndFunc

; ist die Diagnoseerfassung das aktive Fenster? (im Hook, muss schnell sein)
Func DiagAktiv()
	Return StringLeft(_WinAPI_GetWindowText(_WinAPI_GetForegroundWindow()), StringLen($DiagTitel)) = $DiagTitel
EndFunc

; Diagnoseerfassung: klickt die Kurzwahl Nr. $i (0-basiert) in der rechten Fensterhaelfte an
Func DiagnoseWahl($i)
	If $i < 0 Then Return Meldung("Unbekannte Diagnose-Kurzwahl")
	Local $hWnd = WinGetHandle($DiagFenster)
	If @error Then Return Meldung("Diagnoseerfassung nicht geoeffnet")
	If Not WinActive($hWnd) Then
		WinActivate($hWnd)
		If Not WinWaitActive($hWnd, "", 2) Then Return False
		Sleep(100) ; bis das Fenster neu gezeichnet ist
	EndIf
	; bei eingeblendeten Nummern deren Knoepfe nehmen, die Schildchen verdecken sonst die Suche
	Local $k = ($gEinblendung And $gEinblendungWnd = $hWnd) ? $gEinblendungKnoepfe : 0
	EinblendungWeg()
	If Not IsArray($k) Then $k = DiagKnoepfe($hWnd)
	If Not IsArray($k) Then Return Meldung("Keine Diagnose-Kurzwahl gefunden")
	If $i >= UBound($k) Then Return Meldung("Nur " & UBound($k) & " Diagnose-Kurzwahlen sichtbar")
	Return KlickPunkt($hWnd, $k[$i][0], $k[$i][1], "Diagnose-Kurzwahl verdeckt oder nicht sichtbar")
EndFunc

; sucht die Kurzwahl-Knoepfe am Bildschirm und liefert [n][4] in Bildschirmkoordinaten (Mitte x, y,
; Sortierschluessel, linker Rand), geordnet nach [Diagnosen] Reihenfolge, oder 0.
; Die Knoepfe stehen in Spalten dicht untereinander:
; zuerst werden in den ersten Knopfzeilen lange knopfgraue Strecken gesucht (textfreie Pixelzeilen
; eines Knopfs), deren Enden die Spalten ergeben; dann wird je Spalte kurz vor dem rechten Rand
; senkrecht nach den Luecken in der Hintergrundfarbe getastet. Der blaue Fokusrahmen fuellt die
; Luecken um seinen Knopf, solche zusammenhaengenden Stuecke werden nach dem Zeilenabstand geteilt.
Func DiagKnoepfe($hWnd)
	Local $pt = DllStructCreate("int X;int Y")
	_WinAPI_ClientToScreen($hWnd, $pt)
	Local $gr = WinGetClientSize($hWnd)
	If @error Then Return 0
	Local $farbe = Int(IniRead($Ini, "Diagnosen", "KnopfFarbe", $defDiagKnopfFarbe))
	Local $tol = Int(IniRead($Ini, "Diagnosen", "Toleranz", $defDiagToleranz))
	Local $x1 = $pt.X + Int($gr[0] / 2), $x2 = $pt.X + $gr[0] - 1, $y1 = $pt.Y, $y2 = $pt.Y + $gr[1] - 1
	If Not BildLesen($x1, $y1, $x2, $y2) Then Return 0
	; Spalten [n][2]: linker und rechter Rand; nur Strecken ab ca. 100 Pixel, so lang ist in der Mitte
	; zentrierter Text nie auf einer Seite frei
	Local $spalten[0][2], $oben = -1
	For $y = $y1 To $y2 Step 2
		If $oben >= 0 And $y > $oben + 60 Then ExitLoop
		Local $n = 0
		For $x = $x1 To $x2 + 5 Step 5
			If FarbeNah($x, $y, $farbe, $tol) Then
				$n += 1
				ContinueLoop
			EndIf
			If $n >= 20 Then
				Local $e = $x - 5
				While FarbeNah($e + 1, $y, $farbe, $tol)
					$e += 1
				WEnd
				SpalteMerken($spalten, $x - 5 * $n, $e)
				If $oben < 0 Then $oben = $y
			EndIf
			$n = 0
		Next
	Next
	If UBound($spalten) = 0 Then
		$gBild = 0
		Return 0
	EndIf
	Local $zeilen = (IniRead($Ini, "Diagnosen", "Reihenfolge", $defDiagReihenfolge) = "Zeilen")
	Local $k[0][4], $m = 0
	For $c = 0 To UBound($spalten) - 1
		Local $li = $spalten[$c][0], $re = $spalten[$c][1], $mx = Int(($li + $re) / 2)
		; Hintergrund rechts neben der Spalte
		Local $hg = BildFarbe($re + 2, $oben)
		If $hg < 0 Then ContinueLoop
		Local $st[0], $lg[0], $start = -1
		For $y = $oben To $y2 + 1
			Local $luecke = ($y > $y2) Or (FarbeNah($re - 1, $y, $hg, $tol) And FarbeNah($re - 3, $y, $hg, $tol))
			If Not $luecke And $start < 0 Then $start = $y
			If $luecke And $start >= 0 Then
				If $y - $start >= 8 Then
					ReDim $st[UBound($st) + 1], $lg[UBound($lg) + 1]
					$st[UBound($st) - 1] = $start
					$lg[UBound($lg) - 1] = $y - $start
				EndIf
				$start = -1
			EndIf
		Next
		If UBound($st) = 0 Then ContinueLoop
		; typische Knopfhoehe und Luecke
		Local $h = Median($lg), $abst[UBound($st)], $na = 0
		For $j = 1 To UBound($st) - 1
			$abst[$na] = $st[$j] - $st[$j - 1] - $lg[$j - 1]
			$na += 1
		Next
		ReDim $abst[$na]
		Local $l = ($na > 0) ? Median($abst) : 2
		For $j = 0 To UBound($st) - 1
			Local $anz = Round(($lg[$j] + $l) / ($h + $l))
			ReDim $k[$m + ($anz > 1 ? $anz : 1)][4]
			If $anz <= 1 Then
				$k[$m][0] = $mx
				$k[$m][1] = $st[$j] + Int($lg[$j] / 2)
				$k[$m][3] = $li
				$m += 1
			Else
				For $z = 0 To $anz - 1
					$k[$m][0] = $mx
					$k[$m][1] = $st[$j] + $z * ($h + $l) + Int($h / 2)
					$k[$m][3] = $li
					$m += 1
				Next
			EndIf
		Next
	Next
	$gBild = 0
	If $m = 0 Then Return 0
	; Sortierschluessel: spaltenweise ergibt sich die Reihenfolge schon so, zeilenweise nach y, dann x
	For $j = 0 To $m - 1
		$k[$j][2] = $zeilen ? Round($k[$j][1] / 4) * 100000 + $k[$j][0] : $j
	Next
	For $j = 1 To $m - 1
		Local $kx = $k[$j][0], $ky = $k[$j][1], $ks = $k[$j][2], $kl = $k[$j][3], $q = $j
		While $q > 0 And $k[$q - 1][2] > $ks
			For $s = 0 To 3
				$k[$q][$s] = $k[$q - 1][$s]
			Next
			$q -= 1
		WEnd
		$k[$q][0] = $kx
		$k[$q][1] = $ky
		$k[$q][2] = $ks
		$k[$q][3] = $kl
	Next
	Return $k
EndFunc

; nimmt eine Strecke $a..$e als Spalte auf, falls nicht schon eine mit (fast) demselben rechten Rand
; da ist, und haelt die Spalten nach links sortiert
Func SpalteMerken(ByRef $spalten, $a, $e)
	Local $n = UBound($spalten)
	For $i = 0 To $n - 1
		If Abs($spalten[$i][1] - $e) <= 3 Then
			If $a < $spalten[$i][0] Then $spalten[$i][0] = $a
			If $e > $spalten[$i][1] Then $spalten[$i][1] = $e
			Return
		EndIf
	Next
	ReDim $spalten[$n + 1][2]
	While $n > 0 And $spalten[$n - 1][1] > $e
		$spalten[$n][0] = $spalten[$n - 1][0]
		$spalten[$n][1] = $spalten[$n - 1][1]
		$n -= 1
	WEnd
	$spalten[$n][0] = $a
	$spalten[$n][1] = $e
EndFunc

; Median eines Zahlenfelds
Func Median($a)
	For $i = 1 To UBound($a) - 1
		Local $v = $a[$i], $j = $i
		While $j > 0 And $a[$j - 1] > $v
			$a[$j] = $a[$j - 1]
			$j -= 1
		WEnd
		$a[$j] = $v
	Next
	Return $a[Int(UBound($a) / 2)]
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
	Return FarbeGleich(BildFarbe($x, $y), $farbe, $tol)
EndFunc

; weicht die Farbe $c je Farbanteil hoechstens $tol von $farbe ab? ($c = -1: nein)
Func FarbeGleich($c, $farbe, $tol)
	If $c < 0 Then Return False
	Return Abs(BitAND(BitShift($c, 16), 255) - BitAND(BitShift($farbe, 16), 255)) <= $tol _
			And Abs(BitAND(BitShift($c, 8), 255) - BitAND(BitShift($farbe, 8), 255)) <= $tol _
			And Abs(BitAND($c, 255) - BitAND($farbe, 255)) <= $tol
EndFunc

; ist die MO-Tagesuebersicht das aktive Fenster? (wird im Tastatur-Hook aufgerufen, muss schnell sein)
Func TagesAktiv()
	Return _WinAPI_GetClassName(_WinAPI_GetForegroundWindow()) = $TagesKlasse
EndFunc

; Tagesuebersicht: in der Bereichsauswahl links den Eintrag ueber ($richtung = -1) bzw. unter (1) dem
; gelb umrandeten anklicken, am Ende wieder am anderen Ende beginnen; danach, sobald sich die Liste
; rechts neu aufgebaut hat, deren oberste Zeile anklicken.
; Die Auswahl wird am Bildschirm gesucht: gelber Rahmen links (senkrecht) und rechts in derselben Zeile;
; von dort aus werden im Raster der Eintraege nach oben und unten die weiteren gezaehlt, bis an der
; grauen Randspalte keiner mehr ist.
; Geklickt wird mit der echten Maus und erst nach dem Loslassen von Alt: auf einen per Nachricht
; geschickten Klick (ControlClick) beginnt die Bereichsauswahl ein Ziehen (Halteverbots-Mauszeiger),
; statt den Eintrag zu waehlen, und die Liste rechts bekaeme den Tastaturfokus nicht.
Func BereichWechsel($richtung)
	Local $hWnd = WinGetHandle("[ACTIVE]")
	If _WinAPI_GetClassName($hWnd) <> $TagesKlasse Then Return False
	Local $hGrid = ControlGetHandle($hWnd, "", "[CLASS:TNewStringGrid; INSTANCE:1]")
	If @error Then Return Meldung("Liste der Tagesuebersicht nicht gefunden")
	AltMaskieren()
	If Not ModifierLos() Then Return Meldung("Alt bitte loslassen")
	Local $g = WinGetPos($hGrid)
	If @error Then Return False
	Local $pt = DllStructCreate("int X;int Y")
	_WinAPI_ClientToScreen($hWnd, $pt)
	; Bereich links neben der Liste
	Local $x1 = $pt.X, $x2 = $g[0] - 1, $y1 = $g[1], $y2 = $g[1] + $g[3] - 1
	If $x2 - $x1 < 40 Then Return Meldung("Bereichsauswahl links nicht sichtbar")
	Local $farbe = Int(IniRead($Ini, "Tagesuebersicht", "BereichRahmen", $defBereichRahmen))
	Opt("PixelCoordMode", 1)
	If Not BildLesen($x1, $y1, $x2, $y2) Then Return Meldung("Bildschirm nicht lesbar")
	; linke obere Ecke des Rahmens; ein Treffer zaehlt nur mit senkrechtem Rahmen darunter und
	; demselben Rahmen am rechten Rand
	Local $fx = -1, $fy = -1, $rx = -1, $ys = $y1
	While $ys <= $y2 And $fx < 0
		Local $p = PixelSearch($x1, $ys, $x2, $y2, $farbe, 10)
		If @error Then ExitLoop
		If FarbeNah($p[0], $p[1] + 5, $farbe, 10) And FarbeNah($p[0], $p[1] + 10, $farbe, 10) Then
			For $x = $x2 To $p[0] + 20 Step -1
				If FarbeNah($x, $p[1] + 5, $farbe, 10) Then
					$fx = $p[0]
					$fy = $p[1]
					$rx = $x
					ExitLoop
				EndIf
			Next
		EndIf
		$ys = $p[1] + 1
	WEnd
	If $fx < 0 Then
		$gBild = 0
		Return Meldung("Kein gelb umrandeter Eintrag links gefunden")
	EndIf
	; Rahmenhoehe; der Eintrag ist 2 Pixel groesser nach jeder Seite, dazu 1 Pixel Abstand
	Local $fu = $fy
	While $fu < $fy + 60 And FarbeNah($fx, $fu + 1, $farbe, 10)
		$fu += 1
	WEnd
	Local $h = $fu - $fy + 5, $raster = $h + 1, $oben = $fy - 2, $bx = $fx - 1
	; Oberkanten aller Eintraege, der gewaehlte an Stelle $akt
	Local $eintr[1] = [$oben], $akt = 0
	For $r = -1 To 1 Step 2
		Local $t = $oben + $r * $raster, $typ = BereichTyp($bx, $t, $h)
		While $typ <> ""
			If $typ = "E" Then
				If $r < 0 Then
					VorneEinfuegen($eintr, $t)
					$akt += 1
				Else
					ReDim $eintr[UBound($eintr) + 1]
					$eintr[UBound($eintr) - 1] = $t
				EndIf
			EndIf
			$t += $r * $raster
			$typ = BereichTyp($bx, $t, $h)
		WEnd
	Next
	$gBild = 0
	Local $n = UBound($eintr), $ziel = Mod($akt + $richtung + $n, $n)
	; oberste Zeile rechts merken, um zu sehen, wann die Liste neu aufgebaut ist
	Local $zy = $g[1] + $KopfHoehe
	Local $vorher = PixelChecksum($g[0] + 2, $zy, $g[0] + $g[2] - 20, $zy + $ZeilenHoehe - 1)
	Opt("MouseCoordMode", 1)
	Local $alt = MouseGetPos()
	MouseClick("left", Int(($fx + $rx) / 2), $eintr[$ziel] + Int($h / 2), 1, 0)
	Local $t0 = TimerInit(), $max = Int(IniRead($Ini, "Tagesuebersicht", "BereichWarten", $defBereichWarten))
	While TimerDiff($t0) < $max
		Sleep(50)
		If PixelChecksum($g[0] + 2, $zy, $g[0] + $g[2] - 20, $zy + $ZeilenHoehe - 1) <> $vorher Then ExitLoop
	WEnd
	; bis die Zeile fertig gezeichnet ist
	Sleep(100)
	Local $nameX = $g[0] + Int(IniRead($Ini, "Tagesuebersicht", "NameX", $defNameX)), $zm = $zy + Int($ZeilenHoehe / 2)
	; leere Liste: dort ist nur weisse Flaeche
	If PixelGetColor($nameX, $zm) <> 0xFFFFFF Then MouseClick("left", $nameX, $zm, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	Return True
EndFunc

; Art des Platzes mit Oberkante $t in der Bereichsauswahl, geprueft an der grauen Randspalte $x:
; "E" Eintrag, "T" Trennstrich (innen weiss), "" kein Eintrag mehr (liest aus BildLesen)
Func BereichTyp($x, $t, $h)
	If Not FarbeNah($x, $t, $BereichRand, 12) Or Not FarbeNah($x, $t + $h - 1, $BereichRand, 12) Then Return ""
	Return FarbeNah($x, $t + Int($h / 2), 0xFFFFFF, 8) ? "T" : "E"
EndFunc

; fuegt $wert vorne in das Feld $a ein
Func VorneEinfuegen(ByRef $a, $wert)
	ReDim $a[UBound($a) + 1]
	For $i = UBound($a) - 1 To 1 Step -1
		$a[$i] = $a[$i - 1]
	Next
	$a[0] = $wert
EndFunc

; nach einer geschluckten Alt+Taste eine unbelegte Taste (vkE8) nachschicken, damit Windows das
; Loslassen von Alt nicht als "Alt allein" nimmt und in die Menueleiste bzw. das Systemmenue springt
Func AltMaskieren()
	If Not Gedrueckt(0x12) Then Return
	DllCall("user32.dll", "none", "keybd_event", "byte", 0xE8, "byte", 0, "dword", 0, "ulong_ptr", 0)
	DllCall("user32.dll", "none", "keybd_event", "byte", 0xE8, "byte", 0, "dword", 2, "ulong_ptr", 0)
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
	If $i < 0 Then Return Meldung("Unbekannter Filter")
	EinblendungWeg()
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
	Return KlickPunkt($hWnd, $p[0] + $x, $p[1] + $y, "Filterliste verdeckt oder nicht sichtbar")
EndFunc

; klickt an den Bildschirmpunkt $x,$y auf das Control, das dort liegt; $verdeckt ist die Meldung,
; wenn dort kein Control von $hWnd liegt
Func KlickPunkt($hWnd, $x, $y, $verdeckt)
	WinActivate($hWnd)
	WinWaitActive($hWnd, "", 2)
	If IniRead($Ini, "Allgemein", "EchteMaus", "0") = "1" Then
		Local $alt = MouseGetPos()
		Opt("MouseCoordMode", 1)
		MouseClick("left", $x, $y, 1, 0)
		MouseMove($alt[0], $alt[1], 0)
		Return True
	EndIf
	Local $pt = DllStructCreate("int X;int Y")
	$pt.X = $x
	$pt.Y = $y
	Local $hZiel = _WinAPI_WindowFromPoint($pt)
	; 2 = GA_ROOT: das Control muss zum Medical-Office-Fenster gehoeren
	If $hZiel = 0 Or _WinAPI_GetAncestor($hZiel, 2) <> $hWnd Then Return Meldung($verdeckt)
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
