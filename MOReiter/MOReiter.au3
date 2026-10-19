#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=MOReiter.ico
#AutoIt3Wrapper_Outfile=MOReiter.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Kompilieren mit Icon: Aut2exe.exe /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico
; MOReiter: waehlt in Medical Office den Reiter "Kartei", "Krankenblatt", "ePA" oder "ePAAbr" per Mausklick
; Aufruf:  MOReiter.exe Kartei | Krankenblatt | ePA | ePAAbr  -> einmal klicken und beenden (z.B. aus VB6 per Shell)
;          MOReiter.exe LetztePatienten | Menue  -> Pfeil neben der Patientensuche bzw. drei Striche links oben
;          MOReiter.exe Unten                    -> Liste im Container Karteikarte aktivieren
;          MOReiter.exe Filter n                 -> Kartei und dort den n-ten Filter waehlen
;          MOReiter.exe Diagnose n               -> in der Diagnoseerfassung die n-te Kurzwahl rechts anklicken
;          MOReiter.exe Wechsel                  -> Kartei, bzw. Krankenblatt, wenn Kartei schon aktiv ist
;          MOReiter.exe ePAWechsel               -> ePA, bzw. ePAAbr, wenn ePA schon aktiv ist
;          MOReiter.exe                          -> bleibt resident mit Hotkeys:
;            Strg+Alt+K        Kartei, ist Kartei schon aktiv: Krankenblatt
;            Strg+Alt+L        Krankenblatt
;            Strg+Alt+P        ePA, ist ePA schon aktiv: ePAAbr
;            Strg+Alt+U        ("unten") die Liste im Container Karteikarte aktivieren (Fokus hinein), notfalls
;                              vorher zum Krankenblatt wechseln; [Karteikarte] Art=Klick klickt statt dessen
;                              in die markierte (sonst oberste) Zeile
;            Strg+Alt+T        aus jedem Programm: Termine des Patienten zeigen, solange Strg+Alt+T gehalten
;                              wird (Medical Office nach vorn, F6, Alt+T; beim Loslassen F6 und zurueck ins
;                              vorige Fenster)
;            Strg+Alt+F1       bildschirmfuellende Hilfe (Text aus MOReiter-Hilfe.txt, in die exe eingebettet),
;                              Esc oder nochmals Strg+Alt+F1 schliesst sie (MO-Hilfe weiter mit Strg+F1)
;            Strg+Alt+B        (Zusatz, nur mit [Labor] Url= in der INI, Beispiel Labor Staber)
;                              Laborbefunde des Patienten im Firefox: onlinebefunde.labor-staber.de anwaehlen
;                              oder oeffnen, anmelden (Passwort aus dem Firefox), Patient nach Name und
;                              Geburtsdatum suchen und seinen ersten Befund zeigen; die Arbeit auf der
;                              Seite macht das Benutzerskript MOReiter-Labor.user.js (einmal in das
;                              Firefox-Add-on Tampermonkey aufnehmen); Name und Geburtsdatum aus der
;                              Seite "Personalien"
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
;                              Zeile der Liste (Alt darf fuer mehrere Schritte gehalten bleiben)
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
;            Strg+Alt eine halbe Sekunde halten ([Allgemein] Einblendung=ms, 0 = nie): bis Strg oder
;                              Alt losgelassen wird, stehen gelbe Schildchen an allen Stellen im aktiven
;                              Fenster, auf die eine Taste gerade wirken wuerde:
;                              Hauptfenster: K, L, P an den Reitern (K bzw. P an dem, zu dem gewechselt
;                              wuerde), U an der Karteikarte, Z am Pfeil neben der Patientensuche, Leer am Menueknopf, H am
;                              Hochladen-Pfeil der markierten Zeile, an den Filtern der Kartei die
;                              Ziffernblock-Nummer und dahinter die direkte Taste, wenn sie anders heisst
;                              (z.B. "0 ^", "5", "11 sz"); Diagnoseerfassung: ebenso an den Kurzwahlen;
;                              Tagesuebersicht: A in der markierten Zeile, Alt+Pfeil an den Bereichen;
;                              Briefversand: die Buchstaben an den Klickstellen
;            Die Tasten werden per Tastatur-Hook abgefangen (nicht per HotKeySet), weil nur so linkes Alt
;            von AltGr unterschieden werden kann: AltGr kommt als linke Strg + rechte Alt an und bleibt
;            unberuehrt, so dass AltGr+2 3 7 8 9 0 sz weiterhin hoch2 hoch3 { [ ] } \ liefern.
;            (Das fruehere Durchreichen per Send liess gelegentlich die Strg-Taste haengen.)
;            im Fenster "MEDICAL OFFICE - Briefversand" statt der obigen Belegung:
;            Strg+Alt+K Kontaktverzeichnis, I KIM-Verzeichnis, R Krankenblatt, O Ordner (Anhaenge),
;                              E Kaestchen "Empfangsbestaetigung anfordern", B Betreff-Zeile, M oberster
;                              Empfaenger, F Brief-Vorschau, V Versenden/Drucken; Esc allein: Abbrechen
;                              (geklickt wird mit der echten Maus, Strg+Alt darf dabei gehalten bleiben,
;                              z.B. fuer E und dann V);
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
#include <EditConstants.au3>
Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 4)
; keine 250 ms Pause nach jedem Fensterbefehl; wo ein Fenster Zeit braucht, wird ausdruecklich gewartet
Opt("WinWaitDelay", 0)

Const $MOFenster = "[REGEXPTITLE:^Medical Office; CLASS:OWL_Window]"
; INI im Benutzerprofil, weil das Programmverzeichnis unter "Program Files" nicht beschreibbar ist;
; hier stehen nur die eigenen Einstellungen und Kalibrierungen dieses Benutzers
Const $Ini = @AppDataDir & "\MOReiter\MOReiter.ini"
; Voreinstellungen der Praxis neben der exe (verteilt z.B. NVerb), gelten, wo $Ini nichts angibt
Const $IniVorgaben = @ScriptDir & "\MOReiter-Vorgaben.ini"

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
; Labor (Strg+Alt+B, nur wenn [Labor] Url= gesetzt ist): Seitentitel des Portals, Gesamtzeit, Zahl der
; durchsuchten Tabs
Const $defLaborTitel = "Onlinebefunde", $defLaborZeit = 30000, $defLaborTabs = 5
; Installationsquelle des Benutzerskripts: Tampermonkey bietet beim Oeffnen einer *.user.js-Adresse die Installation an
Const $defLaborSkript = "https://raw.githubusercontent.com/libelle17/vb6wurzel/main/MOReiter/MOReiter-Labor.user.js"
; andere Portale: Benutzerskript MOReiter-<Abschnitt>.user.js daneben
Const $defSkriptBasis = "https://raw.githubusercontent.com/libelle17/vb6wurzel/main/MOReiter/MOReiter-"
Const $defTermineWarten = 200 ; ms zwischen F6 und Alt+T bzw. nach dem zweiten F6
Const $defEinblendung = 500 ; ms Strg+Alt halten, bis die Nummern eingeblendet werden
; Ziffernblock: im Hook gesammelte Nummer, ihr Ziel ("Filter"/"Diagnose"), geschluckte, noch
; gedrueckte Ziffertasten als ",vk,", und die gerade angezeigte Nummer
Global $gNummer = "", $gNummerZiel = "", $gUnten = "", $gNummerAngezeigt = ""
; die Nummer stammt von der oberen Ziffernreihe (dann heisst eine einzelne 0 Nr. 10, wie die Taste 0)
Global $gNummerOben = False
; Kartei: bei gehaltenem Strg+Alt bisher getippte Ziffern der oberen Reihe und Zeit der letzten
Global $gFolge = "", $gFolgeZeit = 0
; Strg links/rechts und Alt links, wie sie der Hook zuletzt gesehen hat (ohne die eigenen, mit
; $Kennung eingespeisten Tasten); gilt statt GetAsyncKeyState, solange MOReiter sie Windows gegenueber
; als losgelassen gemeldet hat ($gLosGemeldet), obwohl sie noch gehalten werden (hoechstens 10 s)
Global $gGehaltenMod[3] = [False, False, False], $gLosGemeldet = False, $gLosZeit = 0
Const $Kennung = 0x4D4F5231 ; "MOR1" in dwExtraInfo eigener Tastenereignisse
Const $defFolgeZeit = 2000 ; ms, eine laengere Pause zwischen zwei Ziffern beginnt eine neue Folge
; Einblendung: Fenster mit den Schildchen, zugehoeriges Fenster und dessen Kurzwahl-Knoepfe,
; Beginn des Haltens von Strg+Alt (0 = nicht gehalten), schon versucht
Global $gEinblendung = 0, $gEinblendungWnd = 0, $gEinblendungKnoepfe = 0, $gHaltStart = 0, $gEinblendungVersucht = False
; Tastatur-Hook: der Rueckruf merkt sich nur die Aufgabe, ausgefuehrt wird sie in der Hauptschleife,
; weil Windows einen Hook, der zu lange braucht, stillschweigend abhaengt
Global $gHook = 0, $gAufgabe = "", $gGehalten = 0, $gGehaltenZeit = 0
; Strg+Alt+Esc: MOReiter neu starten (NeustartPruefen), auch wenn es gerade in einer Aufgabe haengt
Global $gNeustart = False, $gHookAdlib = TimerInit()
; offenes Hilfefenster (0 = keins)
Global $gHilfe = 0
; [Aufrufe]: externe Programme auf Strg[+Umsch]+Buchstabe in Medical Office, je Zeile Name, Taste (vk), Umsch, Befehl
Global $gAufrufe[0][4]
; [Aufrufe] Protokoll=1: jede passende Taste mit Fokus und Entscheidung nach %APPDATA%\MOReiter\Aufrufe.log
; (im Hook nur gesammelt, geschrieben in der Hauptschleife)
Global $gAufrufProtokoll = False, $gProtokoll = "", $gFokusInfo = ""
; Beschreibung des letzten COM-Fehlers (ADODB); ohne Handler wuerde jeder COM-Fehler das Programm beenden
Global $gComFehler = "", $gComHandler = ObjEvent("AutoIt.Error", "ComFehler")
; Patient (Nachname|Vorname|Geburtsdatum), den Strg+S in Smart Pix nur ueber den Nachnamen fand: ein zweites
; Strg+S binnen 10 min traegt dem angeklickten Patienten das Geburtsdatum ein
Global $gSmartPixNachtrag = "", $gSmartPixNachtragZeit = 0
; Text der letzten Warnung von Smart Pix (z.B. "Keine Schreibrechte fuer die Datensaetze"), s. SmartPixWarnungWeg
Global $gSmartPixWarnung = ""
; Text der letzten Meldung (fuer die Abschlussmeldung der Smart Pix-Liste, wo der Hinweis sonst verloren ginge)
Global $gMeldung = ""
; Ablaufprotokoll der Smart Pix-Liste (Pfad, waehrend des Durchgangs; "" = keins)
Global $gSpLog = "", $gSpNochmal = False, $gSpTabs = 0
; Verschiebung (y) der Klickstelle [SmartPix] Dateiname=, mit der das Feld zuletzt lesbar war (s. SmartPixDialog)
Global $gSmartPixDateiDy = 0
; [Portal] Ablage= und Datei=: nach "fertig" kommende PDFs des Patienten aus dem Download-Ordner dorthin
; verschieben; je Auftrag eine Zeile: Portal, Ziel, Beginn (TimerInit), vorher vorhandene PDFs ("|a.pdf|"),
; Name;Groesse;Zeit des Kandidaten, regulaerer Ausdruck fuer den Dateinamen, neuer Name, Zeitraum, Patient
; (Nachname, Vorname, Geburtsdatum durch @LF getrennt, fuer DokEinordnen)
Global $gAblagen[0][9], $gAblagenGeprueft = 0
; eingelesener Bildschirmausschnitt fuer die Pfeilpruefung (BildLesen/BildFarbe)
Global $gBild = 0, $gBildX = 0, $gBildY = 0, $gBildB = 0
; Aufruf ueber einen Link "moreiter:..." (Patientenlaufzettel, s. LinkAufruf): Patient und Markierungen aus dem
; Link statt aus Medical Office
Global $gLink = False, $gLinkNach = "", $gLinkVor = "", $gLinkGeb = "", $gLinkPid = "", $gLinkMark[0]

If $CmdLine[0] > 0 And $CmdLine[1] = "Neustart" Then
	; von NeustartPruefen gestartet: warten, bis sich die alte Instanz beendet hat
	Local $seit = TimerInit()
	While _Singleton("MOReiter_resident", 1) = 0
		If TimerDiff($seit) > 15000 Then Exit
		Sleep(200)
	WEnd
ElseIf $CmdLine[0] > 1 And $CmdLine[1] = "Filter" Then
	Exit FilterWahl(Int($CmdLine[2]) - 1) ? 0 : 1
ElseIf $CmdLine[0] > 1 And $CmdLine[1] = "Diagnose" Then
	Exit DiagnoseWahl(Int($CmdLine[2]) - 1) ? 0 : 1
ElseIf $CmdLine[0] > 0 And StringLeft($CmdLine[1], 9) = "moreiter:" Then
	Exit LinkAufruf($CmdLine[1]) ? 0 : 1
ElseIf $CmdLine[0] > 0 Then
	Exit Ausfuehren($CmdLine[1]) ? 0 : 1
ElseIf _Singleton("MOReiter_resident", 1) = 0 Then ; laeuft schon
	; sonst merkt man nicht, dass noch eine alte Version mit alten Hotkeys aktiv ist
;	MsgBox(64, "MOReiter", "MOReiter laeuft bereits." & @CRLF & "Fuer eine neue Version die alte zuerst ueber das Tray-Menue beenden.", 10)
	Exit
EndIf

; Icon aus der Datei daneben laden, falls es beim Kompilieren nicht in die exe gekommen ist
If FileExists(@ScriptDir & "\MOReiter.ico") Then TraySetIcon(@ScriptDir & "\MOReiter.ico")
TraySetToolTip("MOReiter " & Stand() & ": Strg+Alt+F1 Hilfe, K Kartei/Wechsel, L Krankenblatt, P ePA/ePAAbr, U Karteikarte, T Termine, B Labor, Z letzte Patienten, Leertaste Menue, ^..F12 oder Ziffernblock Filter bzw. Diagnose-Kurzwahl")
; sonst haelt ein Klick aufs Tray-Symbol das Skript an, und Windows haengt den unbeantworteten Hook ab
Opt("TrayAutoPause", 0)
AufrufeLaden()
LinkSchemaEintragen()
Global $gRueckruf = DllCallbackRegister("TastenHook", "lresult", "int;wparam;lparam")
If Not HookErneuern() Then
	MsgBox(16, "MOReiter", "Tastatur-Hook konnte nicht eingerichtet werden.", 10)
	Exit 1
EndIf
OnAutoItExitRegister("HookEntfernen")
; Adlib laeuft auch waehrend der Sleeps in einer gerade laufenden Aufgabe
AdlibRegister("NeustartPruefen", 250)
; Strg+Alt+Esc auch als Hotkey des Systems: haengt MOReiter in einer Aufgabe, hat Windows den Hook evtl.
; schon abgehaengt (dann kaeme Alt+Esc, Fensterwechsel, durch); laeuft der Hook, schluckt er die Taste vorher
HotKeySet("^!{ESC}", "NeustartTaste")
; der Hinweis der alten Instanz verschwindet mit ihr: hier 3 s lang bestaetigen
If $CmdLine[0] > 0 And $CmdLine[1] = "Neustart" Then Hinweis("MOReiter neu gestartet")
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
	If $gHilfe Then HilfePruefen()
	If $gProtokoll <> "" Then ProtokollSchreiben()
	If UBound($gAblagen) And TimerDiff($gAblagenGeprueft) > 2000 Then AblagenPruefen()
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
		; selbst eingespeiste Tasten (ModifierAus, TasteEinspeisen) unbesehen weiter, sonst loeste
		; z.B. das fuer Strg+Alt+T gesendete Alt+T bei noch gehaltenem Strg+Alt erneut Termine aus
		If $kb.dwExtraInfo = $Kennung Then Return _WinAPI_CallNextHookEx($gHook, $nCode, $wParam, $lParam)
		Local $vk = $kb.vkCode
		; Strg/Alt mitverfolgen, eigene eingespeiste Ereignisse nicht
		If $vk >= 0xA2 And $vk <= 0xA4 And $kb.dwExtraInfo <> $Kennung Then
			$gGehaltenMod[$vk - 0xA2] = Not ($wParam = $WM_KEYUP Or $wParam = $WM_SYSKEYUP)
			If Not ($gGehaltenMod[0] Or $gGehaltenMod[1] Or $gGehaltenMod[2]) Then $gLosGemeldet = False
		EndIf
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
			; Strg+Alt+Esc in jedem Programm: neu starten; nicht ueber $gAufgabe, das wartet auf die laufende Aufgabe
			If $vk = 0x1B And StrgAlt() Then
				$gNeustart = True
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
	If ($vk = 0x26 Or $vk = 0x28) And AltAllein() And Not Gedrueckt(0x10) And TagesAktiv() Then Return ($vk = 0x26) ? "BereichHoch" : "BereichRunter"
	; Esc allein im Briefversand: Abbrechen
	If $vk = 0x1B And Not (Gedrueckt(0x10) Or Gedrueckt(0x11) Or Gedrueckt(0x12) Or Gedrueckt(0x5B) Or Gedrueckt(0x5C)) _
			And BriefAktiv() Then Return "Brief:Abbrechen"
	; Strg[+Umsch]+Buchstabe aus [Aufrufe], nur in Medical Office und nicht in einem Eingabefeld
	If UBound($gAufrufe) > 0 And StrgAllein() Then
		For $i = 0 To UBound($gAufrufe) - 1
			If $vk = $gAufrufe[$i][1] And Gedrueckt(0x10) = $gAufrufe[$i][2] Then
				Local $mo = MOAktiv(), $text = $mo ? TextFokus() : False
				If $gAufrufProtokoll Then Protokoll($gAufrufe[$i][0] & ": " & ($mo ? "Medical Office" : "nicht Medical Office (" & _
						_WinAPI_GetClassName(_WinAPI_GetForegroundWindow()) & " """ & _WinAPI_GetWindowText(_WinAPI_GetForegroundWindow()) & """)") & _
						($mo ? ", " & $gFokusInfo : "") & " -> " & (($mo And Not $text) ? "Aufruf" : "durchgelassen"))
				If $mo And Not $text Then Return "Aufruf:" & $i
				; Strg+S auch in der Patientenliste von Smart Pix (Geburtsdatum nachtragen)
				If $gAufrufe[$i][3] = "SmartPix" And SmartPixListeAktiv() Then Return "Aufruf:" & $i
				ExitLoop
			EndIf
		Next
	EndIf
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
		Case 0x55 ; U
			Return $kalib ? "" : "Unten"
		Case 0x54 ; T
			Return $kalib ? "" : "Termine"
		Case 0x42 ; B (im Briefversand schon oben als Betreff vergeben)
			Return $kalib ? "" : "Labor"
		Case 0x70 ; F1
			Return $kalib ? "" : "Hilfe"
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
	If $gLosGemeldet And TimerDiff($gLosZeit) > 10000 Then $gLosGemeldet = False
	If $gLosGemeldet Then Return ($gGehaltenMod[0] Or $gGehaltenMod[1]) And $gGehaltenMod[2] And Not Gedrueckt(0xA5)
	Return (Gedrueckt(0xA2) Or Gedrueckt(0xA3)) And Gedrueckt(0xA4) And Not Gedrueckt(0xA5)
EndFunc

; Strg ohne Alt und Windowstaste
Func StrgAllein()
	Return Gedrueckt(0x11) And Not (Gedrueckt(0x12) Or Gedrueckt(0x5B) Or Gedrueckt(0x5C))
EndFunc

; hat im Vordergrundfenster ein Eingabefeld den Fokus (blinkender Textcursor oder Klasse wie Edit, Memo,
; Combo)? Dort behalten Strg+A, Strg+P usw. ihre gewohnte Bedeutung; $gFokusInfo beschreibt den Fokus
Func TextFokus()
	$gFokusInfo = "Fokus unbekannt"
	Local $gti = DllStructCreate("dword cbSize;dword flags;hwnd hwndActive;hwnd hwndFocus;hwnd hwndCapture;" & _
			"hwnd hwndMenuOwner;hwnd hwndMoveSize;hwnd hwndCaret;long left;long top;long right;long bottom")
	DllStructSetData($gti, "cbSize", DllStructGetSize($gti))
	Local $tid = _WinAPI_GetWindowThreadProcessId(_WinAPI_GetForegroundWindow(), 0)
	Local $r = DllCall("user32.dll", "bool", "GetGUIThreadInfo", "dword", $tid, "struct*", $gti)
	If @error Or Not $r[0] Then Return False
	Local $f = DllStructGetData($gti, "hwndFocus"), $kl = $f ? _WinAPI_GetClassName($f) : ""
	; 1 = GUI_CARETBLINKING; ein bloss angelegter, unsichtbarer Textcursor zaehlt nicht
	Local $cursor = BitAND(DllStructGetData($gti, "flags"), 1) <> 0
	$gFokusInfo = "Fokus " & ($f ? $kl : "keiner") & ($cursor ? ", Textcursor" : "")
	Return $cursor Or ($f And StringRegExp($kl, "(?i)edit|memo|rich|combo"))
EndFunc

; Alt ohne Strg gehalten (wie StrgAlt nach ModifierAus an den selbst mitverfolgten Tasten)
Func AltAllein()
	If $gLosGemeldet And TimerDiff($gLosZeit) > 10000 Then $gLosGemeldet = False
	If $gLosGemeldet Then Return $gGehaltenMod[2] And Not ($gGehaltenMod[0] Or $gGehaltenMod[1])
	Return Gedrueckt(0x12) And Not Gedrueckt(0x11)
EndFunc

; meldet gehaltenes Strg und Alt Windows gegenueber als losgelassen, damit ein folgender echter
; Mausklick ohne sie ankommt, obwohl sie noch gedrueckt sind; der Hook erkennt Strg+Alt+Taste dann
; weiter an den selbst mitverfolgten Tasten. Alt zuerst, solange Strg noch gilt (dann ist es kein
; Alt allein, das die Menueleiste aufruft); das Loslassen der echten Tasten kommt spaeter doppelt an,
; was nichts ausloest.
Func ModifierAus()
	If Not (Gedrueckt(0x11) Or Gedrueckt(0x12)) Then Return
	AltMaskieren()
	$gLosGemeldet = True
	$gLosZeit = TimerInit()
	For $vk In StringSplit("164,165,163,162", ",", 2) ; LAlt, RAlt, RStrg, LStrg
		; 1 = KEYEVENTF_EXTENDEDKEY (rechte Tasten), 2 = KEYEVENTF_KEYUP
		If Gedrueckt(Int($vk)) Then DllCall("user32.dll", "none", "keybd_event", "byte", Int($vk), "byte", 0, _
				"dword", 2 + ((Int($vk) = 163 Or Int($vk) = 165) ? 1 : 0), "ulong_ptr", $Kennung)
	Next
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
			And TimerDiff($gFolgeZeit) < Int(Einst("Allgemein", "FolgeZeit", $defFolgeZeit)) Then
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
	Local $ms = Int(Einst("Allgemein", "Einblendung", $defEinblendung))
	If $ms <= 0 Or $gEinblendungVersucht Or TimerDiff($gHaltStart) < $ms Then Return
	$gEinblendungVersucht = True
	EinblendungZeigen()
EndFunc

; Schildchen "Nummer Taste" links an jeder Diagnose-Kurzwahl bzw. an jedem Filter der Kartei; ein
; einziges durchsichtiges, nicht anklickbares Fenster (Farbschluessel Magenta) ueber dem Zielfenster
Func EinblendungZeigen()
	; $pos: linker Rand des Schildchens, Mitte y, Text, Mitte x (-1: linksbuendig angesetzt)
	Local $hWnd = _WinAPI_GetForegroundWindow(), $pos[0][4]
	If BriefAktiv() Then
		BriefPositionen($hWnd, $pos)
	ElseIf DiagAktiv() Then
		Local $k = DiagKnoepfe($hWnd)
		If Not IsArray($k) Then Return
		$gEinblendungKnoepfe = $k
		ReDim $pos[UBound($k)][4]
		For $i = 0 To UBound($k) - 1
			$pos[$i][0] = $k[$i][3] + 2
			$pos[$i][1] = $k[$i][1]
			$pos[$i][2] = NummerText($i)
			$pos[$i][3] = -1
		Next
	ElseIf TagesAktiv() Then
		TagesPositionen($hWnd, $pos)
	ElseIf MOAktiv() Then
		MOPositionen($hWnd, $pos)
	EndIf
	Local $n = UBound($pos)
	If $n = 0 Then Return
	Local $f = WinGetPos($hWnd)
	If @error Then Return
	Local $g = GUICreate("MOReiter-Einblendung", $f[2], $f[3], $f[0], $f[1], $WS_POPUP, _
			BitOR($WS_EX_LAYERED, $WS_EX_TRANSPARENT, $WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_NOACTIVATE))
	GUISetBkColor(0xFF00FF, $g)
	; 1 = LWA_COLORKEY: Magenta ist durchsichtig
	DllCall("user32.dll", "bool", "SetLayeredWindowAttributes", "hwnd", $g, "dword", 0xFF00FF, "byte", 255, "dword", 1)
	For $i = 0 To $n - 1
		Local $t = $pos[$i][2]
		GUICtrlCreateLabel($t, $pos[$i][0] - $f[0], $pos[$i][1] - $f[1] - 8, SchildBreite($t), 16, BitOR($SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, 0xFFE45C)
		GUICtrlSetColor(-1, 0x000000)
		GUICtrlSetFont(-1, 8.5, 700, 0, "Segoe UI")
	Next
	GUISetState(@SW_SHOWNOACTIVATE, $g)
	$gEinblendung = $g
	$gEinblendungWnd = $hWnd
EndFunc

Func EinblendungWeg()
	If $gEinblendung Then
		GUIDelete($gEinblendung)
		; bis der Bildschirm ohne die Schildchen neu gezeichnet ist, sonst stoeren sie Bildschirmpruefungen
		Sleep(50)
	EndIf
	$gEinblendung = 0
	$gEinblendungWnd = 0
	$gEinblendungKnoepfe = 0
EndFunc

; Beschriftung fuer Filter bzw. Diagnose-Kurzwahl Nr. $i: Ziffernblock-Nummer (ab 0), dahinter die
; direkte Taste, falls sie nicht gleich heisst
Func NummerText($i)
	Local $t = String($i)
	If $i < UBound($TastenNamen) And $TastenNamen[$i] <> $t Then $t &= " " & $TastenNamen[$i]
	Return $t
EndFunc

Func SchildBreite($t)
	Return 7 * StringLen($t) + 8
EndFunc

; haengt ein Schildchen mittig ueber dem Bildschirmpunkt $x,$y an $pos an; wirken zwei Tasten auf
; dieselbe Stelle, kommen sie auf ein gemeinsames Schildchen
Func PosDazu(ByRef $pos, $x, $y, $text)
	Local $n = UBound($pos)
	For $i = 0 To $n - 1
		If $pos[$i][3] >= 0 And Abs($pos[$i][3] - $x) <= 4 And Abs($pos[$i][1] - $y) <= 4 Then
			$pos[$i][2] &= " " & $text
			$pos[$i][0] = $pos[$i][3] - Int(SchildBreite($pos[$i][2]) / 2)
			Return
		EndIf
	Next
	ReDim $pos[$n + 1][4]
	$pos[$n][0] = $x - Int(SchildBreite($text) / 2)
	$pos[$n][1] = $y
	$pos[$n][2] = $text
	$pos[$n][3] = $x
EndFunc

; Medical Office Hauptfenster: K, L, P an den Reitern, auf die sie jetzt wirken, Z am Pfeil neben
; der Patientensuche, Leer am Menueknopf, H am Hochladen-Pfeil, dazu die Filter der Kartei
Func MOPositionen($hWnd, ByRef $pos)
	Local $hCtrl = ControlGetHandle($hWnd, "", Einst("Allgemein", "Control", $defCtrl))
	If Not @error And BitAND(WinGetState($hCtrl), 2) Then
		Local $p = WinGetPos($hCtrl), $x, $y
		If Not @error Then
			HolPos(ReiterAktiv("Kartei", $hWnd, $hCtrl) ? "Krankenblatt" : "Kartei", $x, $y)
			PosDazu($pos, $p[0] + $x, $p[1] + $y, "K")
			HolPos("Krankenblatt", $x, $y)
			PosDazu($pos, $p[0] + $x, $p[1] + $y, "L")
			HolPos(ReiterAktiv("ePA", $hWnd, $hCtrl) ? "ePAAbr" : "ePA", $x, $y)
			PosDazu($pos, $p[0] + $x, $p[1] + $y, "P")
		EndIf
	EndIf
	KnopfPos($pos, LetztePatKnopf($hWnd), "Z")
	KnopfPos($pos, MenueKnopf($hWnd), "Leer")
	Local $hGrid = KarteiListe($hWnd), $hx, $hy
	If $hGrid And HochladenStelle($hGrid, $hx, $hy) = "" Then PosDazu($pos, $hx, $hy, "H")
	; U mitten auf die Kopfzeile der Karteikarte
	Local $g = $hGrid ? WinGetPos($hGrid) : 0
	If IsArray($g) Then PosDazu($pos, $g[0] + Int($g[2] / 2), $g[1] + Int($KbKopfHoehe / 2), "U")
	FilterPositionen($hWnd, $pos)
EndFunc

; Schildchen mitten auf den sichtbaren Knopf $h
Func KnopfPos(ByRef $pos, $h, $text)
	If Not $h Or Not BitAND(WinGetState($h), 2) Then Return
	Local $p = WinGetPos($h)
	If @error Then Return
	PosDazu($pos, $p[0] + Int($p[2] / 2), $p[1] + Int($p[3] / 2), $text)
EndFunc

; Tagesuebersicht: A in der markierten Zeile (Spalte Bereich, nur bei "Offene To-Dos") und
; Alt+Pfeil an den Eintraegen der Bereichsauswahl, die Alt+Pfeil hoch bzw. runter waehlen wuerde
Func TagesPositionen($hWnd, ByRef $pos)
	Local $hGrid = ControlGetHandle($hWnd, "", "[CLASS:TNewStringGrid; INSTANCE:1]")
	If @error Then Return
	Local $g = WinGetPos($hGrid)
	If @error Then Return
	Opt("PixelCoordMode", 1)
	If OffeneToDos($hWnd) Then
		Local $mitte = MarkierteZeile($g)
		If $mitte >= 0 Then PosDazu($pos, $g[0] + Int(Einst("Tagesuebersicht", "BereichX", $defBereichX)), $mitte, "A")
	EndIf
	Local $eintr, $akt, $mx, $h
	If BereichSuchen($hWnd, $hGrid, $eintr, $akt, $mx, $h) <> "" Then Return
	Local $n = UBound($eintr)
	If $n < 2 Then Return
	PosDazu($pos, $mx, $eintr[Mod($akt - 1 + $n, $n)] + Int($h / 2), "Alt+" & ChrW(0x2191))
	PosDazu($pos, $mx, $eintr[Mod($akt + 1, $n)] + Int($h / 2), "Alt+" & ChrW(0x2193))
EndFunc

; haengt Bildschirmpunkte (linker Rand der Filterliste, Mitte des Filters) aller sichtbaren Filter an
; $pos an; nur wenn die Kartei (bzw. [Filter] Reiter) schon aktiv ist
Func FilterPositionen($hWnd, ByRef $pos)
	Local $hCtrl = ControlGetHandle($hWnd, "", Einst("Allgemein", "Control", $defCtrl))
	If @error Or Not BitAND(WinGetState($hCtrl), 2) Then Return 0
	If Not ReiterAktiv(Einst("Filter", "Reiter", $defFilterReiter), $hWnd, $hCtrl) Then Return 0
	Local $p = WinGetPos($hCtrl)
	If @error Then Return 0
	Local $x = $p[0] + Int(Einst("Filter", "X", $defFilterX))
	Local $y0 = $p[1] + Int(Einst("Filter", "Y", $defFilterY))
	Local $abst = Number(Einst("Filter", "Abstand", $defFilterAbstand))
	If $abst < 5 Then Return 0
	Local $pt = DllStructCreate("int X;int Y"), $hListe = 0, $links = 0, $unten = 0, $n = 0, $m = UBound($pos)
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
		ReDim $pos[$m + $n + 1][4]
		$pos[$m + $n][0] = $links
		$pos[$m + $n][1] = $y
		$pos[$m + $n][2] = NummerText($n)
		$pos[$m + $n][3] = -1
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
	If $aufgabe = "Unten" Then Return Unten()
	If $aufgabe = "Termine" Then Return Termine()
	If $aufgabe = "Hilfe" Then Return Hilfe()
	If $aufgabe = "Labor" Then Return Labor()
	If StringLeft($aufgabe, 7) = "Aufruf:" Then Return Aufruf(Int(StringTrimLeft($aufgabe, 7)))
	If StringLeft($aufgabe, 7) = "Portal:" Then Return Portal(StringTrimLeft($aufgabe, 7))
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
; Melden von Strg und Alt als losgelassen (sonst kaeme ein Strg+Alt+Klick an); so koennen, ohne Strg+Alt
; loszulassen, mehrere Tasten nacheinander folgen, z.B. E und V
Func BriefKlick($was)
	Local $hWnd = _WinAPI_GetForegroundWindow()
	If StringLeft(_WinAPI_GetWindowText($hWnd), StringLen($BriefTitel)) <> $BriefTitel Then Return False
	Local $x, $y
	If Not BriefZiel($hWnd, $was, $x, $y) Then Return Meldung($was & " im Briefversand nicht gefunden (Strg+Alt+Umsch+D listet die Controls)")
	ModifierAus()
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
	Local $x, $y
	For $i = 0 To UBound($tasten) - 1
		If BriefZiel($hWnd, $tasten[$i][0], $x, $y) Then PosDazu($pos, $x, $y, $tasten[$i][1])
	Next
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
				$x = $p[0] + Int(Einst("Briefversand", "EmpfaengerX", $defEmpfaengerX))
				$y = $p[1] + Int(Einst("Briefversand", "EmpfaengerY", $defEmpfaengerY))
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
	Local $ctl = Einst("Briefversand", $was, "")
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
		Local $nr = ObjCreate("Scripting.Dictionary"), $hFokus = ControlGetHandle($hWnd, "", ControlGetFocus($hWnd))
		TextFokus()
		$txt &= $gFokusInfo & @CRLF
		For $i = 1 To $liste[0][0]
			Local $h = $liste[$i][0], $kl = $liste[$i][1]
			$nr.Item($kl) = $nr.Item($kl) + 1
			Local $p = WinGetPos($h)
			If @error Then ContinueLoop
			$txt &= "[CLASS:" & $kl & "; INSTANCE:" & $nr.Item($kl) & "]" _
					& (BitAND(WinGetState($h), 2) ? "" : " (unsichtbar)") & (($h = $hFokus) ? " (Fokus)" : "") _
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
	Local $farbe = Int(Einst("Diagnosen", "KnopfFarbe", $defDiagKnopfFarbe))
	Local $tol = Int(Einst("Diagnosen", "Toleranz", $defDiagToleranz))
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
	Local $zeilen = (Einst("Diagnosen", "Reihenfolge", $defDiagReihenfolge) = "Zeilen")
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
	Local $x, $y
	Local $fehler = HochladenStelle($hGrid, $x, $y)
	If $fehler <> "" Then Return Meldung($fehler)
	Opt("MouseCoordMode", 1)
	Local $alt = MouseGetPos()
	MouseClick("left", $x, $y, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	; im Fenster "MEDICAL OFFICE - ePA3.0 - Datensatz" kurz nach dem Erscheinen "Hochladen" druecken
	Local $hDlg = WinWait($EpaDialog, "", 15)
	If Not $hDlg Then Return Meldung("Fenster ""MEDICAL OFFICE - ePA"" ist nicht erschienen")
	Local $hKnopf = KnopfBereit($hDlg, "Hochladen", 5000)
	If Not $hKnopf Then Return Meldung("Knopf ""Hochladen"" nicht bereit")
	Sleep(Int(Einst("Hochladen", "Warten", $defHochladenWarten)))
	; schon offene Dialoge merken, damit nur ein neu erscheinender Leistungsdialog bestaetigt wird
	Local $vorher = WinList($LeistungDialog)
	For $i = 1 To $vorher[0][0]
		$vorher[$i][0] = BitAND(WinGetState($vorher[$i][1]), 2) ? "sichtbar" : ""
	Next
	If Not KnopfKlick($hDlg, $hKnopf) Then Return False
	; manchmal folgt nach dem Hochladen ein Dialog zur Leistungsdokumentation: dort "Uebernehmen"
	If Not WinWaitClose($hDlg, "", 30) Then Return True
	Local $t = TimerInit(), $max = Int(Einst("Hochladen", "LeistungWarten", $defLeistungWarten))
	While TimerDiff($t) < $max
		Local $liste = WinList($LeistungDialog)
		For $i = 1 To $liste[0][0]
			If Not BitAND(WinGetState($liste[$i][1]), 2) Or InWinList($vorher, $liste[$i][1]) Then ContinueLoop
			Local $hUeb = KnopfBereit($liste[$i][1], ChrW(220) & "bernehmen", 1000)
			If $hUeb Then
				Sleep(Int(Einst("Hochladen", "Warten", $defHochladenWarten)))
				Return KnopfKlick($liste[$i][1], $hUeb)
			EndIf
		Next
		Sleep(50)
	WEnd
	Return True
EndFunc

; sucht in der markierten Zeile der Liste $hGrid den hellblauen Pfeil nach oben und liefert seine
; Bildschirmstelle in $x,$y; Rueckgabe "" bei Erfolg, sonst der Grund
Func HochladenStelle($hGrid, ByRef $x, ByRef $y)
	Local $g = WinGetPos($hGrid)
	If @error Then Return "Krankenblatt-Liste nicht sichtbar"
	Opt("PixelCoordMode", 1)
	Local $p = PixelSearch($g[0] + 4, $g[1] + $KbKopfHoehe, $g[0] + 4, $g[1] + $g[3] - 3, $KbMarkiertFarbe, 6)
	If @error Then Return "Keine markierte Zeile sichtbar"
	$y = $p[1] + $KbPfeilY
	; den Streifen um die Suchhoehe einmal einlesen, statt jedes Pixel einzeln (je ca. 15-20 ms) abzufragen
	If Not BildLesen($g[0] + 2, $y - 24, $g[0] + $g[2] - 18, $y + 6) Then Return "Bildschirm nicht lesbar"
	$x = PfeilSchaft($g[0] + 4, $g[0] + $g[2] - 20, $y, $KbPfeilHell, True)
	Local $dunkel = ($x < 0) ? PfeilSchaft($g[0] + 4, $g[0] + $g[2] - 20, $y, $KbPfeilDunkel, False) : -1
	$gBild = 0
	If $x >= 0 Then Return ""
	If $dunkel >= 0 Then Return "Schon in der ePA (Pfeil ist schwarz)"
	Return "Kein Pfeil zum Hochladen in der markierten Zeile"
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

; Wert aus der eigenen INI, sonst aus den Praxis-Vorgaben, sonst $def; geschrieben (Kalibrieren) wird
; immer nur in die eigene INI, so dass ein neu verteiltes MOReiter-Vorgaben.ini sie nicht ueberschreibt
; [Aufrufe] aus der eigenen INI und MOReiter-Vorgaben.ini (die eigene geht vor) nach $gAufrufe; Schluessel
; "Strg+P" bzw. "Strg+Umsch+P" (nur Buchstaben), Wert die Befehlszeile, {pat} wird zu "Nachname;Vorname;TT.MM.JJJJ"
Func AufrufeLaden()
	Local $n = 0, $quelle[2] = [$Ini, $IniVorgaben]
	ReDim $gAufrufe[0][4]
	For $datei In $quelle
		Local $a = IniReadSection($datei, "Aufrufe")
		If @error Then ContinueLoop
		For $i = 1 To $a[0][0]
			Local $m = StringRegExp($a[$i][0], "(?i)^Strg\+(Umsch\+)?([A-Z])$", 1)
			If @error Or StringStripWS($a[$i][1], 3) = "" Then ContinueLoop
			Local $vk = Asc(StringUpper($m[1])), $umsch = ($m[0] <> "")
			Local $doppelt = False
			For $j = 0 To $n - 1
				If $gAufrufe[$j][1] = $vk And $gAufrufe[$j][2] = $umsch Then $doppelt = True
			Next
			If $doppelt Then ContinueLoop
			ReDim $gAufrufe[$n + 1][4]
			$gAufrufe[$n][0] = ($umsch ? "Strg+Umsch+" : "Strg+") & StringUpper($m[1])
			$gAufrufe[$n][1] = $vk
			$gAufrufe[$n][2] = $umsch
			$gAufrufe[$n][3] = StringStripWS($a[$i][1], 3)
			$n += 1
		Next
	Next
	$gAufrufProtokoll = (Einst("Aufrufe", "Protokoll", "0") = "1")
	If $gAufrufProtokoll Then Protokoll("MOReiter gestartet, " & $n & " Aufrufe eingerichtet")
EndFunc

Func Protokoll($txt)
	$gProtokoll &= @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "  " & $txt & @CRLF
EndFunc

Func ProtokollSchreiben()
	DirCreate(@AppDataDir & "\MOReiter")
	; 1 = anhaengen, 128 = UTF-8 mit BOM
	Local $fh = FileOpen(@AppDataDir & "\MOReiter\Aufrufe.log", 1 + 128)
	If $fh <> -1 Then
		FileWrite($fh, $gProtokoll)
		FileClose($fh)
	EndIf
	$gProtokoll = ""
EndFunc

; Strg[+Umsch]+Buchstabe aus [Aufrufe]: Programm mit den Daten des geoeffneten Patienten starten
Func Aufruf($i)
	If $i < 0 Or $i >= UBound($gAufrufe) Then Return False
	Local $hMO = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $befehl = $gAufrufe[$i][3]
	; Strg+Z=Zucker: alle Portale bzw. Programme nach den Zucker-Markierungen des Patienten
	If $befehl = "Zucker" Then Return Zucker()
	Return AufrufBefehl($gAufrufe[$i][0], $befehl)
EndFunc

; fuehrt einen Befehl aus [Aufrufe] bzw. [Zucker] aus: "Portal:Abschnitt" oder eine Befehlszeile mit {pat}
Func AufrufBefehl($name, $befehl)
	Local $hMO = WinGetHandle($MOFenster)
	If @error And Not $gLink Then Return MOStarten()
	; Portal im Firefox mit eigenem Benutzerskript, z.B. Strg+L=Portal:LibreView
	If StringLeft($befehl, 7) = "Portal:" Then Return Portal(StringStripWS(StringTrimLeft($befehl, 7), 3))
	; Accu-Chek Smart Pix Software: Patient in der Patientenliste suchen und waehlen
	If $befehl = "SmartPix" Then Return SmartPix()
	If $befehl = "SmartPixListe" Then Return SmartPixListe()
	If $befehl = "Skripte" Then Return SkripteEinrichten()
	If StringInStr($befehl, "{pat}") Then
		; F6 (falls die Personalien fehlen) darf nicht als Strg+F6 ankommen
		If Not ModifierLos() Then Return Meldung("Strg bitte loslassen")
		Local $nach, $vor, $geb
		If Not PatientDaten($hMO, $nach, $vor, $geb) Then Return Meldung("Name und Geburtsdatum des Patienten nicht gefunden (Seite Personalien)")
		$befehl = StringReplace($befehl, "{pat}", '"' & $nach & ";" & $vor & ";" & $geb & '"')
	EndIf
	Local $pid = Run($befehl), $fehler = @error
	If $gAufrufProtokoll Then Protokoll($name & ": " & $befehl & ($fehler ? " -> nicht gestartet" : " -> gestartet, Prozess " & $pid))
	If $fehler Then Return Meldung($name & ": Programm nicht gefunden: " & $befehl)
	Return True
EndFunc

; Strg+Z ([Aufrufe] Strg+Z=Zucker): Markierungen des geoeffneten Patienten aus der Datenbank von Medical Office
; lesen und fuer jede, deren Text einen Schluessel aus [Zucker] enthaelt (z.B. Libre=Portal:LibreView, auch
; "Libre Lesegeraet"), dessen Befehl ausfuehren: Portal:Abschnitt, eine Befehlszeile oder eine Taste aus
; [Aufrufe] wie Strg+S; mehrere nacheinander, jeden Befehl nur einmal
Func Zucker()
	Local $hMO = WinGetHandle($MOFenster)
	If @error And Not $gLink Then Return MOStarten()
	If Not ModifierLos() Then Return Meldung("Strg bitte loslassen")
	Local $nach, $vor, $geb
	If Not PatientDaten($hMO, $nach, $vor, $geb) Then Return Meldung("Name und Geburtsdatum des Patienten nicht gefunden (Seite Personalien)")
	; vom Laufzettel: die dort angeklickte Markierung bzw. bei Alt+Z alle
	Local $mark = $gLinkMark
	If Not $gLink Then
		HinweisStehend("Markierungen von " & $vor & " " & $nach & " werden gelesen ...")
		$mark = MOMarkierungen($nach, $vor, $geb)
		If @error Then Return Meldung($mark)
	EndIf
	If UBound($mark) = 0 Then Return Meldung($vor & " " & $nach & ": keine Markierungen in Medical Office")
	; Schluessel aus [Zucker], eigene INI vor den Vorgaben, in der Reihenfolge der Datei
	Local $regeln[0][2], $quelle[2] = [$Ini, $IniVorgaben]
	For $datei In $quelle
		Local $a = IniReadSection($datei, "Zucker")
		If @error Then ContinueLoop
		For $i = 1 To $a[0][0]
			Local $k = StringStripWS($a[$i][0], 3), $neu = True
			For $j = 0 To UBound($regeln) - 1
				If $regeln[$j][0] = $k Then $neu = False
			Next
			If Not $neu Then ContinueLoop
			ReDim $regeln[UBound($regeln) + 1][2]
			$regeln[UBound($regeln) - 1][0] = $k
			$regeln[UBound($regeln) - 1][1] = StringStripWS($a[$i][1], 3)
		Next
	Next
	; je Markierung die erste passende Regel; leerer Befehl = bewusst kein Abruf
	Local $befehle[0][2], $ohne = ""
	For $m In $mark
		Local $treffer = -1
		For $j = 0 To UBound($regeln) - 1
			If StringInStr($m, $regeln[$j][0]) Then
				$treffer = $j
				ExitLoop
			EndIf
		Next
		If $treffer < 0 Then ContinueLoop
		If $regeln[$treffer][1] = "" Then
			$ohne &= ", " & $m
			ContinueLoop
		EndIf
		Local $da = False
		For $j = 0 To UBound($befehle) - 1
			If $befehle[$j][1] = $regeln[$treffer][1] Then $da = True
		Next
		If $da Then ContinueLoop
		ReDim $befehle[UBound($befehle) + 1][2]
		$befehle[UBound($befehle) - 1][0] = $m
		$befehle[UBound($befehle) - 1][1] = $regeln[$treffer][1]
	Next
	If UBound($befehle) = 0 Then Return Meldung($vor & " " & $nach & ": keine abrufbare Zucker-Markierung" & _
			(($ohne = "") ? "" : " (ohne Abruf: " & StringTrimLeft($ohne, 2) & ")"))
	Local $n = UBound($befehle)
	For $j = 0 To $n - 1
		Local $b = $befehle[$j][1], $name = ($gLink ? "Laufzettel " : "Strg+Z ") & ($j + 1) & "/" & $n & " (" & $befehle[$j][0] & ")"
		Hinweis($name & ": " & $b)
		Sleep(1000)
		; Verweis auf eine Taste aus [Aufrufe], z.B. Strg+S fuer Smart Pix
		If StringRegExp($b, "(?i)^Strg\+(Umsch\+)?[A-Z]$") Then
			Local $gef = False
			For $i = 0 To UBound($gAufrufe) - 1
				If StringUpper($gAufrufe[$i][0]) = StringUpper($b) Then
					AufrufBefehl($name, $gAufrufe[$i][3])
					$gef = True
					ExitLoop
				EndIf
			Next
			If Not $gef Then Meldung($name & ": " & $b & " ist in [Aufrufe] nicht eingerichtet")
		Else
			AufrufBefehl($name, $b)
		EndIf
		; die Meldung des vorigen Abrufs soll lesbar bleiben
		If $j < $n - 1 Then Sleep(2000)
	Next
	If $ohne <> "" Then Hinweis("Ohne Abruf: " & StringTrimLeft($ohne, 2))
	Return True
EndFunc

; [Aufrufe] Strg+S=SmartPix: Accu-Chek Smart Pix Software ([SmartPix] Programm=) starten bzw. nach vorn holen,
; "Patientenliste oeffnen" anklicken, im Feld "Suchen" das Geburtsdatum eingeben und, wenn genau ein Patient
; uebrig bleibt, ihn per Doppelklick waehlen; sonst bleibt die Liste (nach Geburtsdatum, ersatzweise nach
; Nachname gefiltert) zum Waehlen stehen. Das Qt-Programm hat keine Controls, deshalb wird an festen Stellen
; relativ zum Fensterinneren geklickt ([SmartPix] Knopf=, Suche=, Zeile= als x,y; negatives y = vom unteren Rand)
; und am Bildschirm nach dunkler Schrift in der Spalte Nachname gesehen
; Ohne Treffer beim Geburtsdatum wird nach dem Nachnamen gefiltert; den richtigen Patienten anklicken und nochmal
; Strg+S (auch in der Patientenliste): dann wird ihm das Geburtsdatum eingetragen, wenn der Name passt und es fehlt
Func SmartPix()
	Local $hMO = WinGetHandle($MOFenster)
	If @error And Not $gLink Then Return MOStarten()
	If Not ModifierLos() Then Return Meldung("Strg bitte loslassen")
	Local $nach, $vor, $geb
	If Not PatientDaten($hMO, $nach, $vor, $geb) Then Return Meldung("Name und Geburtsdatum des Patienten nicht gefunden (Seite Personalien)")
	; zweites Strg+S nach einer Suche nach dem Nachnamen: der in der Liste angeklickte Patient bekommt das Geburtsdatum
	Local $nachtrag = ($gSmartPixNachtrag = $nach & "|" & $vor & "|" & $geb And TimerDiff($gSmartPixNachtragZeit) < 600000)
	$gSmartPixNachtrag = ""
	Local $hListe = SmartPixListeOeffnen()
	Local $warDa = @extended
	If Not $hListe Then Return False
	If $nachtrag And $warDa Then Return SmartPixNachtragen($hListe, $nach, $vor, $geb)
	Local $n = SmartPixSuche($hListe, $geb), $perName = False
	If $n = 0 Then
		$n = SmartPixSuche($hListe, $nach)
		$perName = True
	EndIf
	If $n < 0 Then Return False
	; weder Geburtsdatum noch Nachname: nach einer zweiten, langsameren Suche neu anlegen
	If $n = 0 Then
		Sleep(1500)
		$n = SmartPixSuche($hListe, $geb, 2)
		$perName = False
		If $n = 0 Then
			$n = SmartPixSuche($hListe, $nach, 2)
			$perName = True
		EndIf
		If $n < 0 Then Return False
		If $n = 0 Then Return SmartPixNeu($hListe, $nach, $vor, $geb)
	EndIf
	; nur andere Patienten mit diesem Nachnamen, alle mit (anderem) Geburtsdatum: der Gesuchte fehlt, also anlegen
	If $perName And Not SmartPixOhneGeb($hListe) Then Return SmartPixNeu($hListe, $nach, $vor, $geb)
	If $n = 1 And Not $perName Then
		If Not SmartPixKlick($hListe, Einst("SmartPix", "Zeile", "60,88"), 2) Then Return False
		If Not SmartPixWeg($hListe, 10000) Then Return Meldung("Smart Pix: Patient nicht gewaehlt, bitte selbst ""Auswaehlen""")
		Hinweis("Smart Pix: " & $vor & " " & $nach & ", " & $geb)
		Return True
	EndIf
	If $perName Then
		$gSmartPixNachtrag = $nach & "|" & $vor & "|" & $geb
		$gSmartPixNachtragZeit = TimerInit()
		HinweisStehend("Smart Pix: kein Eintrag mit " & $geb & ", " & ($n = 1 ? "ein" : "mehrere") & " mit Nachname " & $nach & _
				": bitte " & $vor & " " & $nach & " anklicken und nochmal Strg+S, dann wird das Geburtsdatum dort eingetragen" & _
				" (oder selbst waehlen)")
		AdlibRegister("HinweisWeg", 15000)
		Return True
	EndIf
	HinweisStehend("Smart Pix: mehrere Patienten mit " & $geb & ": bitte " & $vor & " " & $nach & " selbst waehlen")
	AdlibRegister("HinweisWeg", 8000)
	Return True
EndFunc

; startet Smart Pix bzw. holt es nach vorn und oeffnet die Patientenliste; Handle der aktiven Liste (@extended = 1,
; wenn sie schon offen war) oder 0 nach einer Meldung
Func SmartPixListeOeffnen()
	Local $haupt = "[TITLE:Accu-Chek Smart Pix Software; CLASS:Qt5QWindowIcon]"
	Local $liste = "[TITLE:Patientenliste; CLASS:Qt5QWindowIcon]"
	; Qt legt zu jedem Titel auch unsichtbare Fenster an, daher nur sichtbare zaehlen
	Local $hListe = SmartPixFenster($liste), $warDa = ($hListe <> 0)
	If Not $hListe Then
		Local $hHaupt = SmartPixFenster($haupt)
		If Not $hHaupt And SmartPixFenster($haupt, False) Then
			; nach dem Schliessen laeuft Smart Pix versteckt im Infobereich weiter und laesst sich nicht mehr
			; zeigen (nur leerer Rahmen): dann beenden und neu starten; nur Smart Pix selbst, nicht etwa das
			; "Accu-Chek Smart Pix Connection Framework" im Infobereich
			HinweisStehend("Smart Pix laeuft versteckt im Infobereich, wird neu gestartet ...")
			Local $pl = ProcessList(SmartPixExe())
			For $i = 1 To $pl[0][0]
				ProcessClose($pl[$i][1])
				ProcessWaitClose($pl[$i][1], 10)
			Next
			If SmartPixFenster($haupt, False) Then Return Meldung("Smart Pix laesst sich nicht beenden, bitte ueber das Symbol im Infobereich beenden")
			Sleep(1000)
		EndIf
		If Not $hHaupt Then
			; nach hartem Beenden (oder Absturz) haengen die Sperren dieses PCs noch in palock.xml
			SmartPixSperrenAufraeumen()
			Local $exe = Einst("SmartPix", "Programm", "C:\AccuChek\SPIXSW\Accu-Chek Smart Pix Software.exe")
			If Not FileExists($exe) Then Return Meldung("Smart Pix nicht gefunden: " & $exe & " ([SmartPix] Programm=)")
			HinweisStehend("Smart Pix wird gestartet ...")
			Run('"' & $exe & '"', StringLeft($exe, StringInStr($exe, "\", 0, -1) - 1))
			If @error Then Return Meldung("Smart Pix liess sich nicht starten: " & $exe)
			Local $t = TimerInit()
			While Not $hHaupt And TimerDiff($t) < 90000
				SmartPixInfoWeg()
				Sleep(500)
				$hHaupt = SmartPixFenster($haupt)
			WEnd
			If Not $hHaupt Then Return Meldung("Smart Pix: Hauptfenster nach 90 s nicht da")
			Sleep(2000)
		EndIf
		; beim Start kann noch "Wichtige Informationen" kommen und den Fokus nehmen: bis 10 s versuchen
		Local $t2 = TimerInit()
		Do
			SmartPixInfoWeg()
			WinActivate($hHaupt)
			If WinWaitActive($hHaupt, "", 1) Then ExitLoop
		Until TimerDiff($t2) > 10000
		If Not WinActive($hHaupt) Then Return Meldung("Smart Pix laesst sich nicht nach vorn holen")
		Sleep(500)
		HinweisStehend("Smart Pix: Patientenliste wird geoeffnet ...")
		If Not SmartPixKlick($hHaupt, Einst("SmartPix", "Knopf", "256,20"), 1) Then Return False
		; die Liste braucht beim ersten Mal etwa 10 s
		Local $t3 = TimerInit(), $warten = Int(Einst("SmartPix", "Warten", "30")) * 1000
		While Not $hListe And TimerDiff($t3) < $warten
			Sleep(500)
			$hListe = SmartPixFenster($liste)
		WEnd
		If Not $hListe Then Return Meldung("Smart Pix: Patientenliste nicht erschienen")
	EndIf
	WinActivate($hListe)
	If Not WinWaitActive($hListe, "", 5) Then Return Meldung("Smart Pix: Patientenliste nicht aktiv")
	Sleep(500)
	Return SetExtended($warDa ? 1 : 0, $hListe)
EndFunc

; True, wenn die Patientenliste von Smart Pix das aktive Fenster ist (fuer den Tastatur-Hook, daher ohne Warten)
Func SmartPixListeAktiv()
	Local $h = _WinAPI_GetForegroundWindow()
	Return _WinAPI_GetClassName($h) = "Qt5QWindowIcon" And _WinAPI_GetWindowText($h) = "Patientenliste"
EndFunc

; True, wenn das Fenster innerhalb $ms zu oder unsichtbar ist (Qt versteckt Fenster womoeglich nur)
Func SmartPixWeg($h, $ms)
	Local $t = TimerInit()
	While WinExists($h) And BitAND(WinGetState($h), 2)
		If TimerDiff($t) > $ms Then Return False
		Sleep(250)
	WEnd
	Return True
EndFunc

; zweites Strg+S: dem in der Liste angeklickten Patienten fehlt das Geburtsdatum; nach Pruefung des Namens wird
; es in "Patientendaten bearbeiten" eingetragen, ebenso eine fehlende Patienten-ID (Patientennummer aus Medical
; Office), und der Patient dann ueber die Suche nach dem Geburtsdatum gewaehlt
Func SmartPixNachtragen($hListe, $nach, $vor, $geb)
	; Nummer vorher holen, damit der Dialog nicht waehrend der Datenbankabfrage offen steht
	Local $nrFehler = ""
	Local $nr = MOPatNr($nach, $vor, $geb)
	If @error Then
		$nrFehler = $nr
		$nr = ""
	EndIf
	Local $alt = ClipGet(), $f[4], $fehler = "", $neu = ""
	Local $hDlg = SmartPixDialog($hListe, $f)
	If Not $hDlg Then
		ClipPut($alt)
		Return False
	EndIf
	Local $gebNeu = ($f[2] = "") ? $geb : "", $idNeu = ($f[3] = "") ? $nr : ""
	If Not SmartPixNamePasst($f[0], $f[1], $nach, $vor) Then
		$fehler = "angeklickt ist " & $f[1] & " " & $f[0] & ", nicht " & $vor & " " & $nach
		SmartPixDialogZu($hDlg)
	ElseIf $f[2] <> "" And $f[2] <> $geb Then
		$fehler = $f[1] & " " & $f[0] & " hat dort schon das Geburtsdatum " & $f[2]
		SmartPixDialogZu($hDlg)
	ElseIf $gebNeu = "" And $idNeu = "" Then
		SmartPixDialogZu($hDlg)
	ElseIf Not SmartPixSetzen($hDlg, $gebNeu, $idNeu, $f[1]) Then
		$fehler = "Geburtsdatum bzw. Patienten-ID liessen sich nicht eintragen"
	Else
		$neu = (($gebNeu <> "") ? "Geburtsdatum" : "") & (($gebNeu <> "" And $idNeu <> "") ? " und " : "") & _
				(($idNeu <> "") ? "Patienten-ID " & $idNeu : "") & " eingetragen, "
	EndIf
	ClipPut($alt)
	If $fehler <> "" Then Return Meldung("Smart Pix: " & $fehler)
	If $nrFehler <> "" And $f[3] = "" Then $neu &= "(Patientennummer nicht gefunden: " & $nrFehler & ") "
	; jetzt findet die Suche nach dem Geburtsdatum den Patienten
	If SmartPixSuche($hListe, $geb) = 1 And SmartPixKlick($hListe, Einst("SmartPix", "Zeile", "60,88"), 2) And SmartPixWeg($hListe, 10000) Then
		Hinweis("Smart Pix: " & $neu & $vor & " " & $nach & ", " & $geb & " gewaehlt")
	Else
		Hinweis("Smart Pix: " & $neu & $vor & " " & $nach & ", bitte selbst waehlen")
	EndIf
	Return True
EndFunc

; True, wenn in der gefilterten Liste ein Patient ohne Geburtsdatum steht (dunkle Schrift in der Spalte Nachname,
; keine in der Spalte Geburtsdatum, [SmartPix] GebSpalte= als x von,bis im Fensterinneren) oder wenn die Liste
; bis unten reicht und darunter noch mehr stehen koennten (dann vorsichtshalber auch True)
Func SmartPixOhneGeb($hListe)
	Local $z = SmartPixPunkt($hListe, Einst("SmartPix", "Zeile", "60,88"))
	If @error Then Return True
	Local $unten = SmartPixPunkt($hListe, "0,-125"), $hoehe = Int(Einst("SmartPix", "Zeilenhoehe", "28"))
	Local $sp = StringSplit(Einst("SmartPix", "GebSpalte", "370,455"), ",", 2)
	If UBound($sp) <> 2 Then Return True
	Local $g1 = SmartPixPunkt($hListe, Int($sp[0]) & ",0"), $g2 = SmartPixPunkt($hListe, Int($sp[1]) & ",0"), $y
	Opt("PixelCoordMode", 1)
	For $y = $z[1] To $unten[1] - $hoehe / 2 Step $hoehe
		; letzte ganze Zeile belegt: es koennten weitere Patienten folgen
		If $y + $hoehe > $unten[1] - $hoehe / 2 Then
			PixelSearch($z[0] - 36, $y - 8, $z[0] + 790, $y + 8, 0x606060, 0x60)
			Return Not @error
		EndIf
		PixelSearch($z[0] - 36, $y - 8, $z[0] + 120, $y + 8, 0x303030, 0x30)
		If @error Then ContinueLoop
		PixelSearch($g1[0], $y - 8, $g2[0], $y + 8, 0x303030, 0x30)
		If @error Then Return True
	Next
	Return False
EndFunc

; Patient weder ueber Geburtsdatum noch ueber Nachname gefunden (bzw. nur andere mit Geburtsdatum): ueber
; "Neuer Patient..." anlegen mit Nachname,
; Vorname, Geburtsdatum und (wenn in Medical Office eindeutig) der Patientennummer als Patienten-ID, jedes Feld
; zurueckgelesen; danach ueber die Suche nach dem Geburtsdatum waehlen
Func SmartPixNeu($hListe, $nach, $vor, $geb)
	Local $nrFehler = ""
	Local $nr = MOPatNr($nach, $vor, $geb)
	If @error Then
		$nrFehler = $nr
		$nr = ""
	EndIf
	Local $dlg = "[TITLE:Patientendaten bearbeiten; CLASS:Qt5QWindowIcon]", $alt = ClipGet(), $fehler = ""
	HinweisStehend("Smart Pix: " & $vor & " " & $nach & " nicht in der Patientenliste, wird angelegt ...")
	WinActivate($hListe)
	If Not WinWaitActive($hListe, "", 3) Then Return Meldung("Smart Pix: Patientenliste nicht aktiv")
	If Not SmartPixKlick($hListe, Einst("SmartPix", "Neu", "710,-76"), 1) Then Return False
	Local $t = TimerInit(), $h = 0
	While Not $h And TimerDiff($t) < 10000
		Sleep(200)
		$h = SmartPixFenster($dlg)
	WEnd
	If Not $h Then Return Meldung("Smart Pix: ""Neuer Patient"" nicht erschienen ([SmartPix] Neu=)")
	WinActivate($h)
	If Not WinWaitActive($h, "", 3) Then
		SmartPixDialogZu($h)
		Return Meldung("Smart Pix: ""Neuer Patient"" nicht aktiv")
	EndIf
	Sleep(Int(Einst("SmartPix", "DialogWarten", "500")))
	; leerer Nachname: wirklich der Dialog fuer einen neuen Patienten
	If SmartPixFeld() <> "" Then
		$fehler = "der Dialog ist nicht leer"
	Else
		Send($nach, 1)
		If SmartPixFeld() <> $nach Then $fehler = "Nachname"
	EndIf
	If $fehler = "" Then
		Send("{TAB}")
		Send($vor, 1)
		If SmartPixFeld() <> $vor Then $fehler = "Vorname"
	EndIf
	If $fehler = "" Then
		Send("{TAB}{HOME}")
		Send(StringRegExpReplace($geb, "\D", ""), 1)
		If SmartPixDatum(SmartPixFeld()) <> $geb Then $fehler = "Geburtsdatum"
	EndIf
	If $fehler = "" And $nr <> "" Then
		Send("{TAB}")
		Send($nr, 1)
		If SmartPixFeld() <> $nr Then $fehler = "Patienten-ID"
	EndIf
	If $fehler = "" And Not WinActive($h) Then $fehler = "Dialog nicht mehr aktiv"
	If $fehler = "" Then $fehler = SmartPixBestaetigen($h, ($nr <> "") ? "id" : "geb", $vor)
	ClipPut($alt)
	If $fehler <> "" Then
		SmartPixDialogZu($h)
		Return Meldung("Smart Pix: neuer Patient nicht angelegt (" & $fehler & "), bitte selbst ""Neuer Patient...""")
	EndIf
	Local $info = ($nr = "") ? " (ohne Patienten-ID: " & $nrFehler & ")" : ", Patienten-ID " & $nr
	If SmartPixSuche($hListe, $geb) = 1 And SmartPixKlick($hListe, Einst("SmartPix", "Zeile", "60,88"), 2) And SmartPixWeg($hListe, 10000) Then
		Hinweis("Smart Pix: " & $vor & " " & $nach & ", " & $geb & " neu angelegt und gewaehlt" & $info)
	Else
		Hinweis("Smart Pix: " & $vor & " " & $nach & ", " & $geb & " neu angelegt" & $info & ", bitte selbst waehlen")
	EndIf
	Return True
EndFunc

; oeffnet "Patientendaten bearbeiten" zum markierten Patienten der Liste und liest ueber die Zwischenablage
; Nachname, Vorname, Geburtsdatum (TT.MM.JJJJ, "" wenn leer), Patienten-ID und mit $mitDatei den Dateinamen nach
; $f; Handle des Dialogs (der Cursor steht dann in der Patienten-ID) oder 0 nach einer Meldung
Func SmartPixDialog($hListe, ByRef $f, $mitDatei = False)
	Local $dlg = "[TITLE:Patientendaten bearbeiten; CLASS:Qt5QWindowIcon]"
	SmartPixFreiWarten()
	WinActivate($hListe)
	If Not WinWaitActive($hListe, "", 3) Then Return Meldung("Smart Pix: Patientenliste nicht aktiv")
	If Not SmartPixKlick($hListe, Einst("SmartPix", "Bearbeiten", "849,-76"), 1) Then Return 0
	Local $t = TimerInit(), $h = 0, $w = "", $max = Int(Einst("SmartPix", "DialogMax", "30000"))
	While Not $h And TimerDiff($t) < $max
		Sleep(200)
		$h = SmartPixFenster($dlg)
		If Not $h Then $w = SmartPixWarnungWeg()
		If $w <> "" Then Return Meldung("Smart Pix: " & $w)
	WEnd
	If Not $h Then Return Meldung("Smart Pix: ""Patientendaten bearbeiten"" nicht erschienen ([SmartPix] Bearbeiten=)")
	WinActivate($h)
	If Not WinWaitActive($h, "", 3) Then
		SmartPixDialogZu($h)
		Return Meldung("Smart Pix: ""Patientendaten bearbeiten"" nicht aktiv")
	EndIf
	Sleep(Int(Einst("SmartPix", "DialogWarten", "500")))
	; beim Oeffnen steht der Cursor im Nachnamen, mit Tab folgen Vorname, Geburtsdatum, Patienten-ID
	$f[0] = SmartPixFeld()
	Send("{TAB}")
	$f[1] = SmartPixFeld()
	Send("{TAB}")
	$f[2] = SmartPixDatum(SmartPixFeld())
	Send("{TAB}")
	$f[3] = SmartPixFeld()
	; Dateiname des Tagebuchs (nur lesbar, fuer das letzte Auslesen aus palist.xml): danach steht der Cursor dort
	If UBound($f) > 4 Then $f[4] = ""
	; Lage je nach PC (Skalierung) etwas verschieden: sonst 16 bzw. 32 Punkte hoeher versuchen (nicht tiefer, darunter
	; liegen die Spaltenkoepfe der Geraeteliste), die passende Verschiebung gilt dann fuer die folgenden Patienten
	If $mitDatei And UBound($f) > 4 Then
		Local $xy = StringSplit(Einst("SmartPix", "Dateiname", "300,318"), ",", 2), $dys[4] = [$gSmartPixDateiDy, 0, -16, -32], $dy, $versucht = "|"
		If UBound($xy) = 2 Then
			For $dy In $dys
				If StringInStr($versucht, "|" & $dy & "|") Then ContinueLoop
				$versucht &= $dy & "|"
				If Not SmartPixKlick($h, Int($xy[0]) & "," & (Int($xy[1]) + $dy), 1) Then ExitLoop
				$f[4] = SmartPixFeld()
				If StringRegExp($f[4], "^\{[0-9A-Fa-f-]{36}\}\.dia$") Then
					$gSmartPixDateiDy = $dy
					ExitLoop
				EndIf
				$f[4] = ""
			Next
		EndIf
	EndIf
	If $f[0] = "" Or Not WinActive($h) Then
		SmartPixDialogZu($h)
		Return Meldung("Smart Pix: Felder in ""Patientendaten bearbeiten"" nicht lesbar")
	EndIf
	SmartPixLog("geoeffnet: " & SmartPixKurz($f))
	Return $h
EndFunc

; Inhalt des Eingabefelds mit dem Cursor (Strg+A, Strg+C), "" wenn leer
Func SmartPixFeld()
	ClipPut("")
	Send("^a^c")
	Local $t = TimerInit(), $s = ""
	Do
		Sleep(50)
		$s = ClipGet()
	Until $s <> "" Or TimerDiff($t) > 300
	Return StringStripWS($s, 3)
EndFunc

; Geburtsdatum aus dem Feld mit Eingabemaske: "" ohne Ziffern, TT.MM.JJJJ bei 8 Ziffern, sonst wie gelesen
Func SmartPixDatum($s)
	Local $z = StringRegExpReplace($s, "\D", "")
	If $z = "" Then Return ""
	If StringLen($z) = 8 Then Return StringLeft($z, 2) & "." & StringMid($z, 3, 2) & "." & StringRight($z, 4)
	Return $s
EndFunc

; traegt im offenen Dialog (Cursor in der Patienten-ID) das Geburtsdatum ein, prueft es und bestaetigt mit OK;
; sonst wird der Dialog verworfen
Func SmartPixGebSetzen($h, $geb)
	Return SmartPixSetzen($h, $geb, "")
EndFunc

; wie SmartPixGebSetzen, dazu (falls nicht leer) die Patienten-ID $id; $vor = Vorname im Dialog (fuer
; "Dieser Patient besteht bereits!", s. SmartPixBestaetigen)
Func SmartPixSetzen($h, $geb, $id, $vor = "")
	If Not WinActive($h) Then
		SmartPixDialogZu($h)
		Return False
	EndIf
	If $id <> "" Then
		Send("^a")
		Send($id, 1)
		If SmartPixFeld() <> $id Or Not WinActive($h) Then
			SmartPixDialogZu($h)
			Return False
		EndIf
	EndIf
	If $geb <> "" Then
		Send("+{TAB}{HOME}")
		Send(StringRegExpReplace($geb, "\D", ""), 1)
		If SmartPixDatum(SmartPixFeld()) <> $geb Or Not WinActive($h) Then
			SmartPixDialogZu($h)
			Return False
		EndIf
	EndIf
	If SmartPixBestaetigen($h, ($geb <> "") ? "geb" : "id", $vor) = "" Then Return True
	SmartPixDialogZu($h)
	Return False
EndFunc

; OK im Dialog (Enter bzw. [SmartPix] OK=); "" wenn er sich schliesst, sonst der Fehler. Meldet Smart Pix
; "Dieser Patient besteht bereits!" (gleicher Name und Geburtsdatum wie ein anderer Eintrag), wird die Warnung
; weggeklickt und an den Vornamen ein Leerzeichen angehaengt, bis es geht (hoechstens fuenfmal); $feld: Feld mit
; dem Cursor ("geb" oder "id"), $vor: Vorname im Dialog (leer: kein Anhaengen)
Func SmartPixBestaetigen($h, $feld, $vor)
	Local $ok = Einst("SmartPix", "OK", ""), $w, $t, $palist = Einst("SmartPixListe", "Palist", "P:\Datenbanken\AccuChek\Diaries\palist.xml")
	Local $stand = SmartPixDateiStand($palist)
	For $n = 0 To 5
		If $ok = "" Then
			Send("{ENTER}")
		ElseIf Not SmartPixKlick($h, $ok, 1) Then
			Return "OK"
		EndIf
		$w = ""
		$t = TimerInit()
		While TimerDiff($t) < 5000
			Sleep(250)
			If Not WinExists($h) Or Not BitAND(WinGetState($h), 2) Then
				SmartPixFreiWarten(True)
				If Not SmartPixGeschrieben($palist, $stand) Then
					SmartPixLog("Dialog zu, aber palist.xml unveraendert: nicht gespeichert")
					$gSmartPixWarnung = "Speichern nicht angekommen (palist.xml unveraendert)"
					Return "Smart Pix: " & $gSmartPixWarnung
				EndIf
				SmartPixLog("gespeichert")
				Return ""
			EndIf
			$w = SmartPixWarnungWeg()
			If $w <> "" Then ExitLoop
		WEnd
		If $w = "" Then
			SmartPixLog("Speichern: Dialog bleibt offen")
			Return "OK"
		EndIf
		SmartPixLog("Speichern gescheitert: " & $w)
		If Not StringInStr($w, "besteht bereits") Or $vor = "" Or $n = 5 Then Return "Smart Pix: " & $w
		; zurueck in den Vornamen (Tab-Reihenfolge Nachname, Vorname, Geburtsdatum, Patienten-ID) und ans Ende
		WinActivate($h)
		If Not WinWaitActive($h, "", 2) Then Return "Dialog nicht mehr aktiv"
		If $feld = "geb" Then Send("+{TAB}")
		If $feld = "id" Then Send("+{TAB 2}")
		$feld = "vor"
		If SmartPixFeld() <> StringStripWS($vor, 3) Then Return "Vorname nicht wiedergefunden"
		Send("{END} ")
	Next
	Return "OK"
EndFunc

; Aenderungszeit und Groesse einer Datei, zum Erkennen, ob Smart Pix sie geschrieben hat
Func SmartPixDateiStand($datei)
	Return FileGetTime($datei, 0, 1) & "|" & FileGetSize($datei)
EndFunc

; hat Smart Pix palist.xml seit $stand neu geschrieben? Der Dialog schliesst sich auch, wenn das Speichern im
; Hintergrund scheitert (dann ist die Ergaenzung nach dem naechsten Neustart weg); bis [SmartPixListe] Pruefen= ms
; (Vorgabe 10000, 0 = nicht pruefen) warten
Func SmartPixGeschrieben($datei, $stand)
	Local $ms = Int(Einst("SmartPixListe", "Pruefen", "10000"))
	If $ms <= 0 Or Not FileExists($datei) Then Return True
	Local $t = TimerInit()
	While SmartPixDateiStand($datei) = $stand
		If TimerDiff($t) > $ms Then Return False
		Sleep(250)
	WEnd
	Return True
EndFunc

; Warnung von Smart Pix ("Accu-Chek Smart Pix Software - Warnung"): Text lesen (Qt-Meldungen kopieren ihn mit
; Strg+C), nach $gSmartPixWarnung merken und mit OK wegklicken; "" wenn keine offen ist
Func SmartPixWarnungWeg()
	Local $h = SmartPixFenster("[REGEXPTITLE:(?i)^Accu-Chek Smart Pix Software.*Warn; CLASS:Qt5QWindowIcon]")
	If Not $h Then Return ""
	WinActivate($h)
	WinWaitActive($h, "", 2)
	Local $alt = ClipGet()
	ClipPut("")
	Send("^c")
	Sleep(300)
	Local $txt = StringStripWS(StringRegExpReplace(ClipGet(), "(?m)^-+\R?|\R+", " "), 7)
	ClipPut($alt)
	If $txt = "" Then $txt = "Warnung (Text nicht lesbar)"
	If WinActive($h) Then Send("{ENTER}")
	If Not SmartPixWeg($h, 3000) Then WinClose($h)
	$gSmartPixWarnung = $txt
	SmartPixLog("Warnung: " & $txt)
	Return $txt
EndFunc

; eine Zeile ins Ablaufprotokoll der Smart Pix-Liste (%APPDATA%\MOReiter\SmartPixListe.log, [SmartPixListe]
; Protokoll=0 = keins), mit Uhrzeit auf die Millisekunde
Func SmartPixLog($txt)
	If $gSpLog = "" Then Return
	FileWriteLine($gSpLog, @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & " " & $txt)
EndFunc

; Patient fuers Protokoll ohne Namen: Anfangsbuchstaben, Laenge der Namen, ob Geburtsdatum bzw. ID vorhanden
Func SmartPixKurz(ByRef $f)
	Return StringLeft($f[0], 1) & StringLen($f[0]) & "/" & StringLeft($f[1], 1) & StringLen($f[1]) & _
			" geb=" & (($f[2] = "") ? "-" : "ja") & " id=" & (($f[3] = "") ? "-" : "ja")
EndFunc

; verwirft den Dialog (Esc, notfalls Fenster schliessen); eine offene Warnung vorher wegklicken
Func SmartPixDialogZu($h)
	SmartPixWarnungWeg()
	WinActivate($h)
	If WinActive($h) Then Send("{ESC}")
	If Not SmartPixWeg($h, 3000) Then WinClose($h)
	SmartPixFreiWarten(True)
	Return True
EndFunc

; Smart Pix gibt die Sperre eines Patienten nach dem Schliessen des Dialogs erst verzoegert frei; wird er (bzw. der
; naechste) vorher geoeffnet, scheitert das Speichern mit "Kein Schreibzugriff" (von Hand ist man nie so schnell):
; warten, bis palock.xml.lck weg ist und palock.xml keine Patientensperre dieses PCs mehr enthaelt
; ([SmartPixListe] FreiWarten= ms, 0 = nicht warten); False, wenn die Zeit ablief. $nachSchliessen: danach noch
; [SmartPixListe] Nachlauf= ms (Vorgabe 1000, 0 = nicht) warten, das Freigeben allein genuegt nicht immer
Func SmartPixFreiWarten($nachSchliessen = False)
	If $nachSchliessen Then
		Local $ok = SmartPixFreiWarten()
		Sleep(Int(Einst("SmartPixListe", "Nachlauf", "1000")))
		Return $ok
	EndIf
	Local $ms = Int(Einst("SmartPixListe", "FreiWarten", "5000"))
	If $ms <= 0 Then Return True
	Local $datei = SmartPixSperrDatei(), $t = TimerInit()
	Local $ich = StringRegExpReplace(SmartPixIch(), "([\\^$.|?*+()\[\]{}])", "\\$1")
	While TimerDiff($t) < $ms
		If Not FileExists($datei & ".lck") And Not StringRegExp(FileRead($datei), '(?i)<PATIENTLOCK\b[^>]*\bUser="' & $ich & '"') Then Return True
		Sleep(100)
	WEnd
	SmartPixLog("Sperre nach " & $ms & " ms nicht frei")
	Return False
EndFunc

; Name fuer Vergleiche: klein, Umlaute ausgeschrieben, ohne Sterne u.ae. und "zzz" (verstorben) vorn, nur Buchstaben
Func SmartPixNorm($s)
	$s = StringLower(StringStripWS($s, 3))
	$s = StringReplace(StringReplace(StringReplace(StringReplace($s, ChrW(228), "ae"), ChrW(246), "oe"), ChrW(252), "ue"), ChrW(223), "ss")
	$s = StringRegExpReplace($s, "^[*^\s\-_.#+]+", "")
	$s = StringRegExpReplace($s, "^zzz[*^\s\-_.#+]*", "")
	Local $t = StringRegExpReplace($s, "[^\pL]+", " ")
	If @error Then $t = StringRegExpReplace($s, "[^a-z]+", " ")
	Return StringStripWS($t, 7)
EndFunc

; gleicher Nachname und gleicher Vorname bzw. gleicher erster Vorname
Func SmartPixNamePasst($spNach, $spVor, $nach, $vor)
	If SmartPixNorm($spNach) <> SmartPixNorm($nach) Then Return False
	Local $a = SmartPixNorm($spVor), $b = SmartPixNorm($vor)
	Return $a = $b Or StringRegExpReplace($a, " .*", "") = StringRegExpReplace($b, " .*", "")
EndFunc

; tippt $text ins Feld "Suchen" der Patientenliste; Ergebnis: 0 = kein Treffer, 1 = genau ein Patient (nur die
; erste Zeile hat einen Nachnamen, weitere Geraete desselben Patienten stehen ohne Namen darunter), 2 = mehrere,
; -1 = Fehler
Func SmartPixSuche($hListe, $text, $mal = 1)
	If Not SmartPixKlick($hListe, Einst("SmartPix", "Suche", "357,-76"), 1) Then Return -1
	Send("^a")
	Send($text, 1)
	Sleep(Int(Einst("SmartPix", "SuchWarten", "1500")) * $mal)
	Local $z = SmartPixPunkt($hListe, Einst("SmartPix", "Zeile", "60,88"))
	If @error Then Return -1
	Local $unten = SmartPixPunkt($hListe, "0,-125")
	Opt("PixelCoordMode", 1)
	; dunkle Schrift (bis etwa 0x60 je Farbanteil) in der Spalte Nachname
	PixelSearch($z[0] - 36, $z[1] - 8, $z[0] + 120, $z[1] + 8, 0x303030, 0x30)
	If @error Then Return 0
	; weitere Patienten darunter, bis kurz ueber die waagrechte Bildlaufleiste
	PixelSearch($z[0] - 36, $z[1] + 16, $z[0] + 120, $unten[1], 0x303030, 0x30)
	Return @error ? 1 : 2
EndFunc

; "x,y" relativ zum Fensterinneren (negatives x vom rechten, negatives y vom unteren Rand) als Bildschirmpunkt [x, y]
Func SmartPixPunkt($hWnd, $xy)
	Local $m = StringRegExp($xy, "^\s*(-?\d+)\s*,\s*(-?\d+)\s*$", 1)
	If @error Then Return SetError(1, 0, 0)
	Local $pt = DllStructCreate("int X;int Y"), $gr = WinGetClientSize($hWnd)
	If @error Then Return SetError(1, 0, 0)
	_WinAPI_ClientToScreen($hWnd, $pt)
	Local $p[2] = [$pt.X + ((Int($m[0]) < 0) ? $gr[0] + Int($m[0]) : Int($m[0])), $pt.Y + ((Int($m[1]) < 0) ? $gr[1] + Int($m[1]) : Int($m[1]))]
	Return $p
EndFunc

; klickt ($anz = 2: doppelt) mit der echten Maus an "x,y" im Fensterinneren und stellt die Maus zurueck
Func SmartPixKlick($hWnd, $xy, $anz)
	Local $p = SmartPixPunkt($hWnd, $xy)
	If @error Then Return Meldung("Smart Pix: Stelle """ & $xy & """ unlesbar ([SmartPix] in der INI)")
	Local $alt = MouseGetPos()
	Opt("MouseCoordMode", 1)
	MouseClick("left", $p[0], $p[1], $anz, 0)
	MouseMove($alt[0], $alt[1], 0)
	Return True
EndFunc

; Handle des ersten sichtbaren (mit $sichtbar = False: irgendeines) Fensters zu $titel, das zum Prozess von Smart
; Pix gehoert (ein minimiertes wird wiederhergestellt), sonst 0
Func SmartPixFenster($titel, $sichtbar = True)
	Local $fe = WinList($titel), $pl = ProcessList(SmartPixExe()), $dabei
	For $i = 1 To $fe[0][0]
		$dabei = False
		For $j = 1 To $pl[0][0]
			If WinGetProcess($fe[$i][1]) = $pl[$j][1] Then $dabei = True
		Next
		If Not $dabei Then ContinueLoop
		If Not $sichtbar Then Return $fe[$i][1]
		If Not BitAND(WinGetState($fe[$i][1]), 2) Then ContinueLoop
		If BitAND(WinGetState($fe[$i][1]), 16) Then
			WinSetState($fe[$i][1], "", @SW_RESTORE)
			Sleep(1000)
		EndIf
		Return $fe[$i][1]
	Next
	Return 0
EndFunc

; Dateiname des Programms von Smart Pix ([SmartPix] Programm=)
Func SmartPixExe()
	Local $exe = Einst("SmartPix", "Programm", "C:\AccuChek\SPIXSW\Accu-Chek Smart Pix Software.exe")
	Return StringMid($exe, StringInStr($exe, "\", 0, -1) + 1)
EndFunc

; schliesst das Fenster "Wichtige Informationen" (Accu-Chek-Uebersicht), falls es beim Start erscheint
Func SmartPixInfoWeg()
	Local $info = "[TITLE:Wichtige Informationen; REGEXPCLASS:^WindowsForms10\.Window]"
	If WinExists($info) Then WinClose($info)
EndFunc

; [Aufrufe] Strg+Umsch+S=SmartPixListe: geht die ganze Patientenliste von Smart Pix Zeile fuer Zeile durch (je
; "Patientendaten bearbeiten", Felder ueber die Zwischenablage) und schreibt sie nach [SmartPixListe] Datei=.
; Probelauf: nur lesen, mit Abgleich gegen Medical Office (patstamm) und quelle (namen), Nachnamen ohne "zzz"
; der Verstorbenen; fehlt das Geburtsdatum und ist es dort eindeutig, steht es unter Vorschlag mit Eintragen=j.
; Eintragen: traegt die Vorschlaege der Zeilen mit Eintragen=j aus dieser Datei ein (gleiche Nummer und gleicher
; Name, also bei unveraenderter Liste) und schreibt dazu ...-eingetragen.csv. Abbruch mit Esc
Func SmartPixListe()
	If Not ModifierLos() Then Return SmartPixFehler("Strg bitte loslassen")
	Local $datei = Einst("SmartPixListe", "Datei", @MyDocumentsDir & "\SmartPix-Patienten.csv")
	Local $a = MsgBox(3 + 32 + 262144, "MOReiter: Smart Pix-Patientenliste", "Ja: die ganze Patientenliste lesen und mit " & _
			"Medical Office und quelle abgleichen, nach" & @CRLF & $datei & @CRLF & "(danach die Frage, ob gleich eingetragen " & _
			"werden soll)" & @CRLF & @CRLF & _
			"Nein: fehlende Geburtsdaten und Patienten-IDs eintragen, laut Spalte Eintragen=j dieser Datei" & @CRLF & @CRLF & _
			"Das dauert Stunden (etwa 2 s je Patient), den PC solange nicht benutzen; Abbruch mit Esc.")
	If $a <> 6 And $a <> 7 Then Return False
	; Lesen und gleich eintragen: eindeutige Vorschlaege sofort, die CSV entsteht trotzdem (Eintragen=eingetragen)
	Local $sofort = False
	If $a = 6 Then
		$a = MsgBox(3 + 32 + 256 + 262144, "MOReiter: Smart Pix-Patientenliste", "Eindeutige Vorschlaege (Eintragen=j) gleich " & _
				"beim Lesen eintragen?" & @CRLF & @CRLF & "Ja: lesen und gleich eintragen, ohne Rueckfragen; die CSV zeigt " & _
				"danach je Zeile ""eingetragen"" bzw. ""Fehler""" & @CRLF & "Nein: nur Probelauf")
		If $a <> 6 And $a <> 7 Then Return False
		$sofort = ($a = 6)
		$a = 6
	EndIf
	Local $eintragen = ($a = 7), $vorschlag = ObjCreate("Scripting.Dictionary"), $db = ObjCreate("Scripting.Dictionary")
	Local $ausgelesen = ObjCreate("Scripting.Dictionary"), $puids = ObjCreate("Scripting.Dictionary"), $ziel = $datei, $fehlerText
	; endete der letzte Durchgang gleicher Art vor dem Ende der Liste (Abbruch, Esc, Max=), dahinter weitermachen?
	Local $art = $eintragen ? "eintragen" : ($sofort ? "sofort" : "lesen"), $stand = @AppDataDir & "\MOReiter\SmartPixListe-Stand.ini"
	Local $ab[4], $abNr = 0, $abWeiter = 0, $abMarkiert = False, $abBasis = 0
	; Ablaufprotokoll (ohne Namen), waechst bis 2 MB, dann neu
	$gSpLog = ""
	If Einst("SmartPixListe", "Protokoll", "1") <> "0" Then
		$gSpLog = @AppDataDir & "\MOReiter\SmartPixListe.log"
		If FileGetSize($gSpLog) > 2000000 Then FileDelete($gSpLog)
		SmartPixLog("=== " & @MDAY & "." & @MON & "." & @YEAR & " Durchgang " & $art & ", " & @ComputerName)
	EndIf
	; ist die Patientenliste schon offen, auf Wunsch beim dort markierten Patienten anfangen (nicht beim Eintragen
	; aus der CSV, dort gehoeren die Nummern zum Schluessel); die Nummern zaehlen hinter dem letzten Stand weiter
	If Not $eintragen And SmartPixFenster("[TITLE:Patientenliste; CLASS:Qt5QWindowIcon]") Then
		Local $wm = MsgBox(3 + 32 + 256 + 262144, "MOReiter: Smart Pix-Patientenliste", "Die Patientenliste von Smart Pix ist offen." & _
				@CRLF & @CRLF & "Ja: beim dort markierten Patienten anfangen, die CSV-Dateien werden ergaenzt" & @CRLF & _
				"Nein: wie sonst (hinter dem letzten Durchgang bzw. von vorn)")
		If $wm <> 6 And $wm <> 7 Then Return False
		$abMarkiert = ($wm = 6)
	EndIf
	If $abMarkiert Then
		If IniRead($stand, "Stand", "Art", "") = $art Then $abBasis = Int(IniRead($stand, "Stand", "Nr", "0")) + Int(IniRead($stand, "Stand", "Weiter", "0"))
	ElseIf IniRead($stand, "Stand", "Art", "") = $art And Int(IniRead($stand, "Stand", "Nr", "0")) > 0 Then
		For $i = 0 To 3
			$ab[$i] = IniRead($stand, "Stand", "F" & $i, "")
		Next
		Local $w = MsgBox(3 + 32 + 262144, "MOReiter: Smart Pix-Patientenliste", "Der letzte Durchgang (" & IniRead($stand, "Stand", "Zeit", "") & _
				") endete beim " & IniRead($stand, "Stand", "Nr", "") & ". Patienten (" & $ab[1] & " " & $ab[0] & ")." & @CRLF & @CRLF & _
				"Ja: hinter ihm weitermachen, die CSV-Dateien werden ergaenzt" & @CRLF & "Nein: von vorn, die CSV-Dateien werden neu geschrieben")
		If $w <> 6 And $w <> 7 Then Return False
		If $w = 6 Then
			$abNr = Int(IniRead($stand, "Stand", "Nr", "0"))
			$abWeiter = Int(IniRead($stand, "Stand", "Weiter", "0"))
		EndIf
	EndIf
	If $eintragen Then
		If Not SmartPixVorschlaege($datei, $vorschlag) Then Return False
		$ziel = StringRegExpReplace($datei, "(?i)\.csv$", "") & "-eingetragen.csv"
	Else
		HinweisStehend("Smart Pix-Liste: Patienten aus Medical Office und quelle werden gelesen ...")
		; Fehler hier als Fenster, nicht als kurzer Hinweis, damit sie nicht uebersehen werden
		Local $ohne = ""
		Local $dbNamen[2] = ["medoff", "quelle"]
		For $welche In $dbNamen
			$fehlerText = DBPatienten($welche, $db)
			If @error Then
				ToolTip("")
				If $ohne <> "" Then
					MsgBox(16 + 262144, "MOReiter", $fehlerText & @CRLF & @CRLF & "Weder Medical Office noch quelle lesbar, Abbruch.")
					Return False
				EndIf
				If MsgBox(4 + 48 + 262144, "MOReiter", $fehlerText & @CRLF & @CRLF & "Ohne " & $welche & " weitermachen?") <> 6 Then Return False
				$ohne = $welche
			EndIf
		Next
		; ohne quelle: Fallbeginne aus Medical Office ab der Umstellung 2025 gelten als unbekannt, denn ob die
		; frueheren Faelle (aus Turbomed) dort alle stehen, ist nicht gesichert; das ist vorsichtig, weil dann eher
		; mehrere Patienten passen und der Eintrag uebergangen wird
		If $ohne = "quelle" Then
			For $fk In $db.Keys
				If StringLeft($fk, 2) = "F|" And Int($db.Item($fk)) >= Int(Einst("SmartPixListe", "MOAb", "20250101")) Then $db.Remove($fk)
			Next
		EndIf
		; letztes Auslesen je Patient: wer erst danach in die Praxis kam, kommt nicht in Frage
		Local $palist = Einst("SmartPixListe", "Palist", "P:\Datenbanken\AccuChek\Diaries\palist.xml")
		$ausgelesen = SmartPixAuslesen($palist)
		If @error Then
			ToolTip("")
			If MsgBox(4 + 48 + 262144, "MOReiter", $palist & " nicht lesbar ([SmartPixListe] Palist=)" & @CRLF & @CRLF & _
					"Ohne Pruefung des letzten Auslesens weitermachen?") <> 6 Then Return False
		EndIf
		; der Dateiname im Dialog laesst sich nicht lesen: Patienten ueber die Protokolle von Smart Pix zuordnen
		HinweisStehend("Smart Pix-Liste: Protokolle von Smart Pix werden gelesen ...")
		SmartPixProtokolle($palist, $puids)
	EndIf
	Local $hListe = SmartPixListeOeffnen()
	If Not $hListe Then Return False
	ToolTip("")
	; beim markierten Patienten: kein Neustart (die Markierung ginge verloren), nur ein Hinweis
	If $abMarkiert Then
		SmartPixSperren()
		If @extended > 1 Then Return SmartPixFehler("Smart Pix-Liste: eigene Listensperre in palock.xml hochgezaehlt, " & _
				"bitte Smart Pix schliessen und ohne ""beim markierten Patienten"" anfangen")
	EndIf
	; hochgezaehlte eigene Listensperre (Smart Pix frueher hart beendet): dann scheitert jedes Speichern, also neu starten
	; (offenes Smart Pix auf anderen PCs stoert nicht, danach wird nicht mehr gefragt)
	If Not $abMarkiert Then
		SmartPixSperren()
		If @extended > 1 Then
			$hListe = SmartPixNeustart()
			If Not $hListe Then Return False
		EndIf
	EndIf
	If Not $abMarkiert And Not SmartPixListeOben($hListe) Then Return False
	Local $zurueck = Einst("SmartPixListe", "ZurTabelle", "")
	Local $neustarts = Int(Einst("SmartPixListe", "Neustarts", "5")), $neugestartet = 0
	If $neustarts < 1 Then $neustarts = 1
	; ohne jeden Patienten zu oeffnen: die Zellen der Liste mit Strg+C lesen ([SmartPixListe] Zellen=0 = wie frueher)
	If Einst("SmartPixListe", "Zellen", "1") <> "0" Then Return SmartPixListeZellen($hListe, $zurueck, $eintragen, $sofort, $art, _
			$stand, $ab, $abNr, $abMarkiert, $abBasis, $ziel, $vorschlag, $db, $neustarts, $puids, $ausgelesen)
	; scheitert das Wiederfinden (meist "Kein Schreibzugriff" beim Oeffnen des Patienten), Smart Pix neu starten
	; und noch einmal suchen, erst dann fragen
	Local $gefunden = ($abNr = 0)
	For $i = 0 To $neustarts
		If $gefunden Then ExitLoop
		If $i > 0 Then
			$neugestartet += 1
			$hListe = SmartPixNeustart()
			If Not $hListe Or Not SmartPixListeOben($hListe) Then Return False
		EndIf
		$gSmartPixWarnung = ""
		$gefunden = SmartPixAnspringen($hListe, $ab, $zurueck)
		; Ergaenzung nicht gespeichert: dieser Patient noch einmal (mit seiner Nummer), die uebersprungenen danach auch
		If $gefunden And $gSpNochmal Then
			$abNr -= 1
			$abWeiter = 0
		ElseIf $gefunden Then
			$gefunden = SmartPixUeberspringen($hListe, $zurueck, $abWeiter)
		EndIf
		If _IsPressed("1B") Then Return False
	Next
	If Not $gefunden Then
		ToolTip("")
		If MsgBox(4 + 48 + 262144, "MOReiter: Smart Pix-Patientenliste", $ab[1] & " " & $ab[0] & " nicht wiedergefunden" & _
				(($gSmartPixWarnung = "") ? "" : " (Smart Pix: " & $gSmartPixWarnung & ")") & "." & @CRLF & @CRLF & "Von vorn anfangen (die CSV-Dateien werden neu geschrieben)?") <> 6 Then Return False
		$abNr = 0
		$abWeiter = 0
		If Not SmartPixListeOben($hListe) Then Return False
	EndIf
	; beim Weitermachen anhaengen (1), sonst neu schreiben (2)
	Local $modus = ($abNr > 0 Or $gSpNochmal Or ($abMarkiert And FileExists($ziel))) ? 1 : 2
	Local $h = FileOpen($ziel, $modus + 8 + 128)
	If $h = -1 Then Return SmartPixFehler("Smart Pix-Liste: " & $ziel & " laesst sich nicht schreiben ([SmartPixListe] Datei=)")
	Local $kopf = "Nr;Nachname;Vorname;Geburtsdatum;Patienten-ID;Letztes Auslesen;Medical Office;quelle;Vorschlag;Vorschlag-ID;Eintragen"
	If $modus = 2 Then FileWriteLine($h, $kopf)
	; beim Lesen zusaetzlich die Liste zum Nachsehen: weder in Medical Office noch in quelle bzw. mehrere passende
	Local $pruefen = StringRegExpReplace($ziel, "(?i)\.csv$", "") & "-pruefen.csv", $h2 = -1, $grund
	If Not $eintragen Then
		$h2 = FileOpen($pruefen, $modus + 8 + 128)
		If $h2 = -1 Then Return SmartPixFehler("Smart Pix-Liste: " & $pruefen & " laesst sich nicht schreiben")
		If $modus = 2 Then FileWriteLine($h2, $kopf & ";Grund")
	EndIf
	Local $alt = ClipGet(), $max = Int(Einst("SmartPixListe", "Max", "0"))
	Local $pause = Int(Einst("SmartPixListe", "Pause", "200")), $grenze = Int(Einst("SmartPixListe", "Gleich", "30"))
	Local $wiederholt = 0, $mitDatei = (Einst("SmartPixListe", "Dateiname", "0") = "1")
	; beim Weitermachen zaehlt die Nummer weiter (beim Eintragen gehoert sie zum Schluessel aus der CSV)
	Local $nr = $abNr + $abWeiter + $abBasis, $nr0 = $abNr + $abBasis, $gleich = 0, $vorher = "", $f[5], $fehlt = 0, $neu = 0, $fehler = 0, $ende = "Ende der Liste", $hDlg, $k
	Local $imMO, $inQu, $vs, $vid, $geb, $id, $teile, $la, $fa, $ids, $ohneDatei = 0, $mehrdeutig = 0, $ein, $f2[4]
	; hinter dem zuletzt gelesenen uebersprungene Patienten (lassen sich auch nach einem Neustart nicht oeffnen)
	Local $skip = $abWeiter, $uebersprungen = 0, $aufgeben = False, $leer[4]
	Local $t0 = TimerInit()
	While 1
		If _IsPressed("1B") Then
			$ende = "mit Esc abgebrochen"
			ExitLoop
		EndIf
		$gSmartPixWarnung = ""
		$gMeldung = ""
		; ohne Klick ins Feld "Dateiname" (laesst sich nicht kopieren; Klicks daneben koennen Kaestchen umschalten),
		; das letzte Auslesen kommt ueber die Protokolle (SmartPixLetztesAuslesen); [SmartPixListe] Dateiname=1 = doch
		$hDlg = SmartPixDialog($hListe, $f, Not $eintragen And $mitDatei)
		; unbeaufsichtigt: noch einmal versuchen, bevor abgebrochen wird; "Kein Schreibzugriff" bzw. "Keine
		; Schreibrechte" kommt bei schnellem Arbeiten ueber das Netz voruebergehend (Sperrdatei palock.xml.lck noch
		; nicht frei): dann mit wachsender Pause, nach jedem dritten Versuch Smart Pix neu starten, erst nach
		; [SmartPixListe] Neustarts= (Vorgabe 5) vergeblichen Neustarts beim selben Patienten abbrechen
		Local $wdh = 0, $schreib = StringInStr($gSmartPixWarnung, "Schreib"), $verloren = False, $hier = 0
		While Not $hDlg And $wdh < (($schreib Or $verloren) ? 3 * $neustarts : 3)
			$wdh += 1
			$wiederholt += 1
			SmartPixLog("Patient " & ($nr + 1) & ": " & $wdh & ". Wiederholung" & ($verloren ? " (Stelle verloren)" : ""))
			ToolTip("Smart Pix-Liste: Patient " & ($nr + 1) & ", " & $wdh & ". Wiederholung" & (($gSmartPixWarnung = "") ? "" : _
					" (" & $gSmartPixWarnung & ")"), @DesktopWidth - 330, 0, "MOReiter")
			Sleep(3000 * (Mod($wdh - 1, 3) + 1))
			SmartPixInfoWeg()
			SmartPixWarnungWeg()
			; hilft das Warten nicht, Smart Pix neu starten (raeumt die eigenen Sperren auf) und zurueck zum Patienten
			; hinter dem zuletzt gelesenen
			; ist die Liste verschwunden (oder liess sich nicht nach vorn holen), sie neu oeffnen und ebenso zurueck
			Local $weg = Not BitAND(WinGetState($hListe), 2)
			; nach gescheitertem Wiederfinden steht die Liste irgendwo: dann jedes Mal neu starten und neu suchen
			If (Mod($wdh, 3) = 0 And $schreib) Or $weg Or $verloren Then
				; ab dem markierten Patienten, aber noch keiner gelesen: nach einem Neustart unauffindbar
				If $abMarkiert And $vorher = "" Then ExitLoop
				If Not $weg Then $neugestartet += 1
				$hListe = $weg ? SmartPixListeOeffnen() : SmartPixNeustart()
				If Not $hListe Or Not SmartPixListeOben($hListe) Then ExitLoop
				$gSmartPixWarnung = ""
				$gMeldung = ""
				If $vorher <> "" Then
					Local $vab = StringSplit($vorher, "|", 2)
					$verloren = Not SmartPixAnspringen($hListe, $vab, $zurueck)
				ElseIf $abNr > 0 Then
					$verloren = Not SmartPixAnspringen($hListe, $ab, $zurueck)
				EndIf
				; Ergaenzung des zuletzt gelesenen nicht gespeichert: er kommt mit seiner Nummer noch einmal dran
				If Not $verloren And $gSpNochmal Then
					SmartPixLog("Patient " & $nr & " noch einmal")
					$nr -= 1
					$skip = 0
					$hier = -1
				ElseIf Not $verloren Then
					$verloren = Not SmartPixUeberspringen($hListe, $zurueck, $skip)
				EndIf
				If _IsPressed("1B") Then ExitLoop
				If $verloren Then
					If StringInStr($gSmartPixWarnung, "Schreib") Then $schreib = True
					ContinueLoop
				EndIf
				; hat schon ein Neustart bei diesem Patienten nicht geholfen, liegt es an ihm (z.B. auf einem anderen
				; PC geoeffnet, Sperre in palock.xml): ueberspringen, in der Liste vermerken, beim naechsten weiter
				If Not $weg Then $hier += 1
				If $hier >= 2 And $schreib Then
					If Not SmartPixWeiter($hListe, $zurueck) Then
						$verloren = True
						ContinueLoop
					EndIf
					Sleep($pause)
					$nr += 1
					$skip += 1
					$uebersprungen += 1
					SmartPixLog("Patient " & $nr & " uebersprungen")
					$leer[0] = "?"
					$leer[1] = "(nach " & (($vorher <> "") ? StringReplace($vorher, "|", " ") : $ab[0] & " " & $ab[1]) & ")"
					SmartPixZeile($h, $nr, $leer, "", "", "", "", "", "uebersprungen")
					If $h2 <> -1 Then SmartPixZeile($h2, $nr, $leer, "", "", "", "", "", "uebersprungen;laesst sich nicht oeffnen (Smart Pix: " & _
							$gSmartPixWarnung & ")")
					IniWrite($stand, "Stand", "Weiter", $skip)
					If $skip >= 5 Then
						$ende = "abgebrochen, " & $skip & " Patienten hintereinander nicht zu oeffnen (Smart Pix: " & $gSmartPixWarnung & ")"
						$aufgeben = True
						ExitLoop
					EndIf
					$wdh = 0
					$hier = 0
				EndIf
			EndIf
			$gSmartPixWarnung = ""
			$gMeldung = ""
			$hDlg = SmartPixDialog($hListe, $f, Not $eintragen And $mitDatei)
			If StringInStr($gSmartPixWarnung, "Schreib") Then $schreib = True
		WEnd
		If $aufgeben Then ExitLoop
		If Not $hDlg Then
			$ende = "abgebrochen, Dialog nicht lesbar" & (($gSmartPixWarnung = "") ? (($gMeldung = "") ? "" : " (" & $gMeldung & ")") : _
					" (Smart Pix: " & $gSmartPixWarnung & ")")
			ExitLoop
		EndIf
		; weitere Geraete desselben Patienten bzw. das Ende der Liste liefern dieselben Daten
		If $f[0] & "|" & $f[1] & "|" & $f[2] & "|" & $f[3] = $vorher Then
			$gleich += 1
			SmartPixDialogZu($hDlg)
			If $gleich >= $grenze Then ExitLoop
		Else
			$gleich = 0
			$skip = 0
			$vorher = $f[0] & "|" & $f[1] & "|" & $f[2] & "|" & $f[3]
			$nr += 1
			SmartPixLog("Patient " & $nr & ": " & SmartPixKurz($f))
			$k = $nr & "|" & SmartPixNorm($f[0]) & "|" & SmartPixNorm($f[1])
			If $eintragen Then
				$geb = ""
				$id = ""
				If $vorschlag.Exists($k) Then
					; nur was in Smart Pix noch fehlt
					$teile = StringSplit($vorschlag.Item($k), "|", 2)
					If $f[2] = "" Then $geb = $teile[0]
					If $f[3] = "" Then $id = $teile[1]
				EndIf
				If $geb <> "" Or $id <> "" Then
					If SmartPixSetzen($hDlg, $geb, $id, $f[1]) Then
						$neu += 1
						SmartPixGemerkt($f, $vorher, $geb, $id)
						SmartPixZeile($h, $nr, $f, "", "", "", $geb, $id, "eingetragen")
					Else
						$fehler += 1
						SmartPixZeile($h, $nr, $f, "", "", "", $geb, $id, "Fehler")
					EndIf
				Else
					SmartPixDialogZu($hDlg)
				EndIf
			Else
				SmartPixDialogZu($hDlg)
				$imMO = SmartPixAbgleich($db, $f, "M")
				$inQu = SmartPixAbgleich($db, $f, "Q")
				$grund = ""
				If StringLeft($imMO, 2) <> "ja" And StringLeft($inQu, 2) <> "ja" Then
					$fehlt += 1
					$grund = "nicht gefunden"
				EndIf
				$la = ""
				$fa = 0
				If $f[4] <> "" And $ausgelesen.Exists(StringLower($f[4])) Then
					$la = $ausgelesen.Item(StringLower($f[4]))
					If $ausgelesen.Exists("E|" & StringLower($f[4])) Then $fa = Int($ausgelesen.Item("E|" & StringLower($f[4])))
				Else
					$la = SmartPixLetztesAuslesen($puids, $ausgelesen, $f)
					$fa = @extended
					If $la = "" Then $ohneDatei += 1
				EndIf
				; Vorschlag nur bei genau einem passenden Patienten, der vor dem letzten Auslesen schon da war (und
				; nicht erst lange nach dem ersten); mehrere stehen zur Ansicht da, ohne Eintragen=j
				$ids = SmartPixKandidaten($db, $f, $la, $fa)
				$vs = ""
				$vid = ""
				For $i = 0 To UBound($ids) - 1
					If $f[2] = "" And $db.Item("P|" & $ids[$i]) <> "" Then $vs &= ($vs = "" ? "" : ", ") & $db.Item("P|" & $ids[$i])
					If $f[3] = "" Then $vid &= ($vid = "" ? "" : ", ") & $ids[$i]
				Next
				If UBound($ids) > 1 Then $mehrdeutig += 1
				$ein = (UBound($ids) = 1 And $vs & $vid <> "") ? "j" : ""
				If $sofort And $ein = "j" Then
					; neu oeffnen, damit der Cursor wieder in der Patienten-ID steht, und pruefen, dass es derselbe ist
					$ein = "Fehler"
					$hDlg = SmartPixDialog($hListe, $f2)
					If $hDlg Then
						If $f2[0] = $f[0] And $f2[1] = $f[1] And $f2[2] = $f[2] And $f2[3] = $f[3] Then
							If SmartPixSetzen($hDlg, $vs, $vid, $f[1]) Then $ein = "eingetragen"
						Else
							SmartPixDialogZu($hDlg)
						EndIf
					EndIf
					If $ein = "eingetragen" Then
						$neu += 1
						SmartPixGemerkt($f, $vorher, $vs, $vid)
					Else
						$fehler += 1
					EndIf
				EndIf
				$la = ($la = "") ? "" : StringRight($la, 2) & "." & StringMid($la, 5, 2) & "." & StringLeft($la, 4)
				SmartPixZeile($h, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein)
				If UBound($ids) > 1 Then $grund = "mehrere passende: " & SmartPixListeText($ids)
				If $grund <> "" Then SmartPixZeile($h2, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein & ";" & $grund)
			EndIf
			; scheitert das Eintragen wiederholt an fehlenden Schreibrechten (z.B. Sperre eines anderen Smart Pix in
			; palock.xml), scheitert auch jeder weitere Patient
			If (StringInStr($gSmartPixWarnung, "Schreib") Or StringInStr($gSmartPixWarnung, "nicht angekommen")) And $fehler >= 5 Then
				$ende = "abgebrochen, Smart Pix: " & $gSmartPixWarnung & " (Sperre in palock.xml? Smart Pix ueberall schliessen)"
				ExitLoop
			EndIf
			; erreichter Stand, fuer das Weitermachen beim naechsten Mal
			IniWrite($stand, "Stand", "Art", $art)
			IniWrite($stand, "Stand", "Nr", $nr)
			IniWrite($stand, "Stand", "Weiter", 0)
			For $i = 0 To 3
				IniWrite($stand, "Stand", "F" & $i, $f[$i])
			Next
			IniWrite($stand, "Stand", "Zeit", @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN)
			If $max > 0 And $nr - $nr0 >= $max Then
				$ende = "nach [SmartPixListe] Max=" & $max
				ExitLoop
			EndIf
			; fest oben rechts, nicht an der Maus ueber der Liste
			If Mod($nr, 10) = 0 Then ToolTip("Smart Pix-Liste: " & $nr & " Patienten, " & Round(TimerDiff($t0) / 60000) & _
					" min (Esc = Abbruch)", @DesktopWidth - 330, 0, "MOReiter")
		EndIf
		; zurueck in die Tabelle und zum naechsten Patienten
		If Not WinActive($hListe) Then WinActivate($hListe)
		If Not WinWaitActive($hListe, "", 3) Then
			$ende = "abgebrochen, Patientenliste nicht aktiv"
			ExitLoop
		EndIf
		If Not SmartPixWeiter($hListe, $zurueck) Then
			$ende = "abgebrochen, markierte Zeile nicht gefunden (s. %APPDATA%\MOReiter\SmartPixKante.txt und .png)"
			ExitLoop
		EndIf
		Sleep($pause)
	WEnd
	FileClose($h)
	If $h2 <> -1 Then FileClose($h2)
	ClipPut($alt)
	ToolTip("")
	If $nr = $nr0 + 1 And $ende = "Ende der Liste" Then $ende = "die Zeile wechselt nicht, [SmartPixListe] ZurTabelle= pruefen"
	; ganz durch: beim naechsten Mal wieder von vorn
	If $ende = "Ende der Liste" Then FileDelete($stand)
	SmartPixLog("Ende bei Patient " & $nr & ": " & $ende)
	$gSpLog = ""
	MsgBox(64 + 262144, "MOReiter: Smart Pix-Patientenliste", (($nr0 > 0 Or $abMarkiert) ? "ab dem " & ($nr0 + 1) & ". bis zum " : "") & _
			$nr & (($nr0 > 0 Or $abMarkiert) ? "." : "") & " Patienten, " & $ende & ", " & Round(TimerDiff($t0) / 60000) & " min" & _
			(($wiederholt = 0) ? "" : ", " & $wiederholt & " Wiederholungen") & _
			(($neugestartet = 0) ? "" : ", " & $neugestartet & " Neustarts von Smart Pix") & _
			(($uebersprungen = 0) ? "" : ", " & $uebersprungen & " Patienten nicht zu oeffnen (uebersprungen)") & _
			@CRLF & ($eintragen ? $neu & " Patienten ergaenzt, " & $fehler & " Fehler" : $fehlt & _
			" weder in Medical Office noch in quelle, " & $mehrdeutig & " mehrdeutig (uebergangen), " & $ohneDatei & _
			" ohne bekanntes letztes Auslesen (ohne diese Pruefung)" & ($sofort ? ", " & $neu & _
			" Patienten ergaenzt, " & $fehler & " Fehler" : "")) & @CRLF & @CRLF & $ziel & _
			($eintragen ? "" : @CRLF & $pruefen))
	Return True
EndFunc

; Smart Pix-Liste ueber die Zellen der Tabelle, ohne jeden Patienten zu oeffnen: Strg+C kopiert die Zelle mit dem
; (unsichtbaren) Zellcursor, Pfeil rechts/links wechselt die Spalte, Pfeil runter die Zeile. Jedes Geraet ist eine
; Zeile, nur die erste eines Patienten traegt Name, Geburtsdatum und Patienten-ID; das letzte Auslesen steht je
; Geraet da. Geoeffnet wird ein Patient nur noch zum Eintragen ([SmartPixListe] Zellen=0 = wie frueher)
Func SmartPixListeZellen($hListe, $zurueck, $eintragen, $sofort, $art, $stand, ByRef $ab, $abNr, $abMarkiert, $abBasis, $ziel, ByRef $vorschlag, ByRef $db, $neustarts, ByRef $puids, ByRef $ausgelesen)
	Local $sp = SmartPixSpalten()
	If @error Then Return SmartPixFehler("Smart Pix-Liste: [SmartPixListe] Spalten= unlesbar (Spalten von Nachname, Vorname, " & _
			"Geburtsdatum, Patienten-ID und Letztes Auslesen, ab 0 gezaehlt, z.B. 0,1,2,3,6)")
	Local $alt = ClipGet(), $r[5], $p[5], $f[5], $erledigt = False, $nochmal = False
	; Anfang: oben (die Liste ist schon nach oben gestellt), beim markierten bzw. beim gemerkten Patienten
	If $abMarkiert Then
		If Not SmartPixWeiter($hListe, $zurueck, False) Or Not SmartPixZurErstenZeile() Then
			ClipPut($alt)
			Return SmartPixFehler("Smart Pix-Liste: markierter Patient nicht gefunden")
		EndIf
	ElseIf $abNr > 0 Then
		If SmartPixZellenFinden($hListe, $ab, $sp) Then
			; fehlt ihm die gemerkte Ergaenzung (nicht gespeichert), kommt er mit seiner Nummer noch einmal dran,
			; sonst wird er nur ueberlesen
			$nochmal = (@extended = 1)
			If $nochmal Then $abNr -= 1
			$erledigt = Not $nochmal
		Else
			ToolTip("")
			If MsgBox(4 + 48 + 262144, "MOReiter: Smart Pix-Patientenliste", $ab[1] & " " & $ab[0] & " nicht wiedergefunden." & _
					@CRLF & @CRLF & "Von vorn anfangen (die CSV-Dateien werden neu geschrieben)?") <> 6 Then
				ClipPut($alt)
				Return False
			EndIf
			$abNr = 0
			If Not SmartPixListeOben($hListe) Then Return False
		EndIf
	EndIf
	; beim Weitermachen anhaengen (1), sonst neu schreiben (2)
	Local $modus = ($abNr > 0 Or $nochmal Or ($abMarkiert And FileExists($ziel))) ? 1 : 2
	Local $h = FileOpen($ziel, $modus + 8 + 128)
	If $h = -1 Then Return SmartPixFehler("Smart Pix-Liste: " & $ziel & " laesst sich nicht schreiben ([SmartPixListe] Datei=)")
	Local $kopf = "Nr;Nachname;Vorname;Geburtsdatum;Patienten-ID;Letztes Auslesen;Medical Office;quelle;Vorschlag;Vorschlag-ID;Eintragen"
	If $modus = 2 Then FileWriteLine($h, $kopf)
	Local $pruefen = StringRegExpReplace($ziel, "(?i)\.csv$", "") & "-pruefen.csv", $h2 = -1, $grund
	If Not $eintragen Then
		$h2 = FileOpen($pruefen, $modus + 8 + 128)
		If $h2 = -1 Then Return SmartPixFehler("Smart Pix-Liste: " & $pruefen & " laesst sich nicht schreiben")
		If $modus = 2 Then FileWriteLine($h2, $kopf & ";Grund")
	EndIf
	Local $max = Int(Einst("SmartPixListe", "Max", "0")), $grenze = Int(Einst("SmartPixListe", "GleichZeilen", "10"))
	Local $zp = Int(Einst("SmartPixListe", "ZeilePause", "30"))
	Local $nr = $abNr + $abBasis, $nr0 = $nr, $fehlt = 0, $neu = 0, $fehler = 0, $ende = "Ende der Liste", $k, $vorher = ""
	Local $imMO, $inQu, $vs, $vid, $geb, $id, $teile, $la, $fa, $ids, $ohneDatei = 0, $mehrdeutig = 0, $ein, $zeilen, $daten, $n
	Local $pla, $pfa
	Local $gelesen = False, $listeEnde = False, $gleich, $vorig, $warum, $folge = 0, $neugestartet = 0, $neustartFolge = 0
	Local $t0 = TimerInit()
	SmartPixLog("Zellen lesen ab Patient " & ($nr + 1))
	While 1
		If _IsPressed("1B") Then
			$ende = "mit Esc abgebrochen"
			ExitLoop
		EndIf
		If Not WinActive($hListe) Then
			WinActivate($hListe)
			If Not WinWaitActive($hListe, "", 3) Then
				$ende = "abgebrochen, Patientenliste nicht aktiv"
				ExitLoop
			EndIf
		EndIf
		; $r ist die erste Zeile des Patienten (schon gelesen, wenn $gelesen)
		If Not $gelesen And Not SmartPixReiheLesen($r, $sp) Then
			$ende = "abgebrochen, Zelle nicht lesbar oder unerwartet (s. %APPDATA%\MOReiter\SmartPixListe.log, [SmartPixListe] Spalten=)"
			ExitLoop
		EndIf
		$gelesen = False
		If $r[0] = "" Then
			$ende = "abgebrochen, Zeile ohne Nachnamen am Anfang eines Patienten (Zellcursor verloren?)"
			ExitLoop
		EndIf
		$p = $r
		$la = ""
		$fa = 0
		$daten = SmartPixDatumDazu($r[4], $la, $fa)
		$zeilen = 1
		; weitere Geraete: folgende Zeilen ohne Nachnamen; am Ende der Liste bleibt Pfeil runter in der letzten Zeile
		$vorig = SmartPixReiheText($r)
		$gleich = 0
		While 1
			If _IsPressed("1B") Then ExitLoop
			Send("{DOWN}")
			Sleep($zp)
			If Not SmartPixReiheLesen($r, $sp) Then ExitLoop
			If SmartPixReiheText($r) = $vorig Then
				$gleich += 1
				If $gleich >= $grenze Then
					$listeEnde = True
					ExitLoop
				EndIf
				ContinueLoop
			EndIf
			$gleich = 0
			$vorig = SmartPixReiheText($r)
			If $r[0] <> "" Then
				$gelesen = True
				ExitLoop
			EndIf
			$daten += SmartPixDatumDazu($r[4], $la, $fa)
			$zeilen += 1
		WEnd
		If Not $gelesen And Not $listeEnde Then
			$ende = _IsPressed("1B") ? "mit Esc abgebrochen" : "abgebrochen, Zelle nicht lesbar oder unerwartet (s. %APPDATA%\MOReiter\SmartPixListe.log, [SmartPixListe] Spalten=)"
			ExitLoop
		EndIf
		For $i = 0 To 3
			$f[$i] = $p[$i]
		Next
		$f[4] = ""
		; der gemerkte, schon fertige Patient beim Weitermachen
		If $erledigt Then
			$erledigt = False
			$vorher = $f[0] & "|" & $f[1] & "|" & $f[2] & "|" & $f[3]
			If $listeEnde Then ExitLoop
			ContinueLoop
		EndIf
		$nr += 1
		SmartPixLog("Patient " & $nr & ": " & SmartPixKurz($f) & ", " & $zeilen & " Zeilen, " & $daten & " Daten")
		$geb = ""
		$id = ""
		$ein = ""
		$vs = ""
		$vid = ""
		If $eintragen Then
			$k = $nr & "|" & SmartPixNorm($f[0]) & "|" & SmartPixNorm($f[1])
			If $vorschlag.Exists($k) Then
				; nur was in Smart Pix noch fehlt
				$teile = StringSplit($vorschlag.Item($k), "|", 2)
				If $f[2] = "" Then $geb = $teile[0]
				If $f[3] = "" Then $id = $teile[1]
			EndIf
		Else
			$imMO = SmartPixAbgleich($db, $f, "M")
			$inQu = SmartPixAbgleich($db, $f, "Q")
			$grund = ""
			If StringLeft($imMO, 2) <> "ja" And StringLeft($inQu, 2) <> "ja" Then
				$fehlt += 1
				$grund = "nicht gefunden"
			EndIf
			; die Liste zeigt womoeglich nicht alle Geraete (verbundene Zellen): dazu das Auslesen aus palist.xml
			$pla = SmartPixLetztesAuslesen($puids, $ausgelesen, $f)
			$pfa = @extended
			If $pla <> "" And ($la = "" Or Int($pla) > Int($la)) Then $la = $pla
			If $pfa > 0 And ($fa = 0 Or $pfa < $fa) Then $fa = $pfa
			If $la = "" Then $ohneDatei += 1
			; Vorschlag nur bei genau einem passenden Patienten, der vor dem letzten Auslesen schon da war (und
			; nicht erst lange nach dem ersten); mehrere stehen zur Ansicht da, ohne Eintragen=j
			$ids = SmartPixKandidaten($db, $f, $la, $fa)
			For $i = 0 To UBound($ids) - 1
				If $f[2] = "" And $db.Item("P|" & $ids[$i]) <> "" Then $vs &= ($vs = "" ? "" : ", ") & $db.Item("P|" & $ids[$i])
				If $f[3] = "" Then $vid &= ($vid = "" ? "" : ", ") & $ids[$i]
			Next
			If UBound($ids) > 1 Then $mehrdeutig += 1
			$ein = (UBound($ids) = 1 And $vs & $vid <> "") ? "j" : ""
			If $sofort And $ein = "j" Then
				$geb = $vs
				$id = $vid
			EndIf
		EndIf
		; eintragen: zurueck in eine Zeile dieses Patienten (damit ist er markiert), oeffnen, eintragen, danach in die
		; erste Zeile des naechsten
		If $geb <> "" Or $id <> "" Then
			; Pfeil hoch so oft, wie der Patient Zeilen hat (am Ende der Liste eine weniger): in seine erste Zeile
			$n = $listeEnde ? $zeilen - 1 : $zeilen
			If $n > 0 Then Send("{UP " & $n & "}")
			Sleep(200)
			$warum = SmartPixZellenSetzen($hListe, $f, $geb, $id)
			; kein Fehler von Smart Pix: weder Neustart noch Fehlerzaehlung
			If $warum = "schon da" Or StringInStr($warum, "in der Liste nicht gelesen") Then
				$ein = $warum
			ElseIf $warum = "" Then
				$neu += 1
				$folge = 0
				$neustartFolge = 0
				SmartPixGemerkt($f, $vorher, $geb, $id)
				$ein = "eingetragen"
			Else
				SmartPixLog("Patient " & $nr & ": " & $warum)
				$folge += 1
				; scheitert es mehrmals hintereinander, Smart Pix neu starten und diesen Patienten noch einmal
				If $folge >= 3 Then
					$neustartFolge += 1
					If $neustartFolge > $neustarts Then
						$ende = "abgebrochen, " & $folge & " Patienten hintereinander nicht eingetragen, auch nach " & $neustarts & _
								" Neustarts (" & $warum & ")"
						ExitLoop
					EndIf
					$folge = 0
					$neugestartet += 1
					$hListe = SmartPixNeustart()
					If Not $hListe Then
						$ende = "abgebrochen, Smart Pix laesst sich nicht neu starten"
						ExitLoop
					EndIf
					If Not SmartPixZellenFinden($hListe, $f, $sp) Then
						$ende = "abgebrochen, nach dem Neustart von Smart Pix Patient " & $nr & " nicht wiedergefunden"
						ExitLoop
					EndIf
					$nr -= 1
					$gelesen = False
					$listeEnde = False
					ContinueLoop
				EndIf
				$fehler += 1
				$ein = "Fehler"
			EndIf
			If Not $listeEnde And Not SmartPixZellenZurueck($hListe, $zurueck, $f, $r, $sp) Then
				$ende = "abgebrochen, nach dem Eintragen den naechsten Patienten nicht wiedergefunden"
				SmartPixZeilenSchreiben($h, $h2, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein, $grund, $ids, $eintragen, $geb, $id)
				ExitLoop
			EndIf
		EndIf
		SmartPixZeilenSchreiben($h, $h2, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein, $grund, $ids, $eintragen, $geb, $id)
		; erreichter Stand, fuer das Weitermachen beim naechsten Mal
		IniWrite($stand, "Stand", "Art", $art)
		IniWrite($stand, "Stand", "Nr", $nr)
		IniWrite($stand, "Stand", "Weiter", 0)
		For $i = 0 To 3
			IniWrite($stand, "Stand", "F" & $i, $f[$i])
		Next
		IniWrite($stand, "Stand", "Zeit", @MDAY & "." & @MON & "." & @YEAR & " " & @HOUR & ":" & @MIN)
		If $listeEnde Then ExitLoop
		If $max > 0 And $nr - $nr0 >= $max Then
			$ende = "nach [SmartPixListe] Max=" & $max
			ExitLoop
		EndIf
		If Mod($nr, 10) = 0 Then ToolTip("Smart Pix-Liste: " & $nr & " Patienten, " & Round(TimerDiff($t0) / 60000) & _
				" min (Esc = Abbruch)", @DesktopWidth - 330, 0, "MOReiter")
	WEnd
	FileClose($h)
	If $h2 <> -1 Then FileClose($h2)
	ClipPut($alt)
	ToolTip("")
	; ganz durch: beim naechsten Mal wieder von vorn
	If $ende = "Ende der Liste" Then FileDelete($stand)
	SmartPixLog("Ende bei Patient " & $nr & ": " & $ende)
	$gSpLog = ""
	MsgBox(64 + 262144, "MOReiter: Smart Pix-Patientenliste", (($nr0 > 0 Or $abMarkiert) ? "ab dem " & ($nr0 + 1) & ". bis zum " : "") & _
			$nr & (($nr0 > 0 Or $abMarkiert) ? "." : "") & " Patienten, " & $ende & ", " & Round(TimerDiff($t0) / 60000) & " min" & _
			(($neugestartet = 0) ? "" : ", " & $neugestartet & " Neustarts von Smart Pix") & _
			@CRLF & ($eintragen ? $neu & " Patienten ergaenzt, " & $fehler & " Fehler" : $fehlt & _
			" weder in Medical Office noch in quelle, " & $mehrdeutig & " mehrdeutig (uebergangen), " & $ohneDatei & _
			" ohne bekanntes letztes Auslesen (ohne diese Pruefung)" & ($sofort ? ", " & $neu & _
			" Patienten ergaenzt, " & $fehler & " Fehler" : "")) & @CRLF & @CRLF & $ziel & _
			($eintragen ? "" : @CRLF & $pruefen))
	Return True
EndFunc

; die Zeilen eines Patienten in die CSV-Dateien (beim Eintragen nur die eingetragenen bzw. gescheiterten)
Func SmartPixZeilenSchreiben($h, $h2, $nr, ByRef $f, $la, $imMO, $inQu, $vs, $vid, $ein, $grund, ByRef $ids, $eintragen, $geb, $id)
	If $eintragen Then
		If $ein <> "" Then SmartPixZeile($h, $nr, $f, "", "", "", $geb, $id, $ein)
		Return
	EndIf
	$la = ($la = "") ? "" : StringRight($la, 2) & "." & StringMid($la, 5, 2) & "." & StringLeft($la, 4)
	SmartPixZeile($h, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein)
	If UBound($ids) > 1 Then $grund = "mehrere passende: " & SmartPixListeText($ids)
	If $grund <> "" Then SmartPixZeile($h2, $nr, $f, $la, $imMO, $inQu, $vs, $vid, $ein & ";" & $grund)
EndFunc

; Spalten der Liste fuer Nachname, Vorname, Geburtsdatum, Patienten-ID und Letztes Auslesen ([SmartPixListe]
; Spalten=, ab 0 gezaehlt, aufsteigend)
Func SmartPixSpalten()
	Local $s = StringSplit(StringStripWS(Einst("SmartPixListe", "Spalten", "0,1,2,3,6"), 8), ",", 2), $sp[5]
	If UBound($s) <> 5 Then Return SetError(1, 0, 0)
	For $i = 0 To 4
		If Not StringIsDigit($s[$i]) Then Return SetError(1, 0, 0)
		$sp[$i] = Int($s[$i])
		If $i > 0 And $sp[$i] <= $sp[$i - 1] Then Return SetError(1, 0, 0)
	Next
	Return $sp
EndFunc

; Inhalt der Zelle mit dem Zellcursor (Strg+C), "" wenn leer: vorher steht ein Merkzeichen in der Zwischenablage;
; eine leere Zelle ersetzt es sofort durch nichts, bleibt es [SmartPixListe] Kopieren= ms (Vorgabe 300) stehen,
; ging das Strg+C verloren: noch bis zu zweimal, dann @error = 1
Func SmartPixZelle()
	Local $merk = "<MOReiter " & Random(100000, 999999, 1) & ">", $ms = Int(Einst("SmartPixListe", "Kopieren", "300")), $s, $t
	ClipPut($merk)
	For $n = 1 To 3
		Send("^c")
		$t = TimerInit()
		Do
			Sleep(10)
			$s = ClipGet()
			If $s <> $merk Then Return StringStripWS($s, 3)
		Until TimerDiff($t) > $ms * $n
	Next
	Return SetError(1, 0, "")
EndFunc

; eine Zeile der Liste lesen; der Zellcursor steht vorher und nachher in der ersten Spalte. $r[0..3] Nachname,
; Vorname, Geburtsdatum, Patienten-ID (nur in der ersten Zeile eines Patienten), $r[4] Letztes Auslesen (bei
; mehreren Geraeten womoeglich mehrere Daten in einer Zelle); False, wenn ein Datum nicht wie eines aussieht
; (Spalten verschoben?), dann steht die Zelle unkenntlich gemacht im Protokoll
Func SmartPixReiheLesen(ByRef $r, ByRef $sp)
	Local $c, $falsch, $zp = Int(Einst("SmartPixListe", "ZellePause", "50")), $zleer = Int(Einst("SmartPixListe", "LeerPause", "250"))
	; ging ein Strg+C oder eine Pfeiltaste verloren (Datum, wo keins sein kann; meist beim Rollen der Liste), die
	; Zeile noch zweimal, vorher sicher in die erste Spalte; dann abbrechen (False)
	For $n = 1 To 3
		If $n > 1 Then
			Sleep(500 * $n)
			Send("{LEFT " & ($sp[4] + 2) & "}")
			Sleep(200)
		EndIf
		$c = 0
		$falsch = ""
		For $i = 0 To 4
			$r[$i] = ""
			If $i >= 1 And $i <= 3 And $r[0] = "" Then ContinueLoop
			If $sp[$i] > $c Then
				Send("{RIGHT " & ($sp[$i] - $c) & "}")
				Sleep($zp)
			EndIf
			$c = $sp[$i]
			$r[$i] = SmartPixZelle()
			If @error Then $falsch &= ", Strg+C ohne Ergebnis in Spalte " & $sp[$i]
			; Smart Pix liefert gelegentlich eine gefuellte Zelle leer: leer erst, wenn es nach einer Pause so bleibt
			If $r[$i] = "" And $falsch = "" Then
				Sleep($zleer)
				$r[$i] = SmartPixZelle()
				If @error Then $falsch &= ", Strg+C ohne Ergebnis in Spalte " & $sp[$i]
				If $r[$i] <> "" Then SmartPixLog("Spalte " & $sp[$i] & " erst beim zweiten Mal gelesen")
			EndIf
		Next
		If $c > 0 Then Send("{LEFT " & $c & "}")
		If $r[2] <> "" And Not StringRegExp($r[2], "^\d\d\.\d\d\.\d{4}$") Then $falsch &= ", Geburtsdatum unerwartet: " & _
				SmartPixMaske($r[2]) & " (Spalte " & $sp[2] & ")"
		If $r[4] <> "" And Not StringRegExp($r[4], "\d\d\.\d\d\.\d{4}") Then $falsch &= ", Letztes Auslesen unerwartet: " & _
				SmartPixMaske($r[4]) & " (Spalte " & $sp[4] & ")"
		If $falsch = "" Then Return True
		SmartPixLog(StringTrimLeft($falsch, 2) & (($n < 3) ? ", Zeile noch einmal" : ""))
	Next
	Return False
EndFunc

; Zelleninhalt fuers Protokoll ohne Patientendaten: Ziffern als 9, Buchstaben als x, Zeilenwechsel als |
Func SmartPixMaske($s)
	$s = StringRegExpReplace(StringRegExpReplace(StringRegExpReplace($s, "\R", "|"), "\d", "9"), "[^9|.,:;/() -]", "x")
	Return StringLen($s) & " Zeichen: " & StringLeft($s, 120)
EndFunc

Func SmartPixReiheText(ByRef $r)
	Return $r[0] & "|" & $r[1] & "|" & $r[2] & "|" & $r[3] & "|" & $r[4]
EndFunc

; Daten TT.MM.JJJJ der Geraete (eins je Zeile oder mehrere in einer Zelle): $la = spaetestes (JJJJMMTT), $fa =
; fruehestes (Zahl JJJJMMTT); gibt die Zahl der Daten zurueck
Func SmartPixDatumDazu($d, ByRef $la, ByRef $fa)
	Local $a = StringRegExp($d, "(\d\d)\.(\d\d)\.(\d{4})", 4), $j, $m
	If @error Then Return 0
	For $i = 0 To UBound($a) - 1
		$m = $a[$i]
		$j = $m[3] & $m[2] & $m[1]
		If $la = "" Or Int($j) > Int($la) Then $la = $j
		If $fa = 0 Or Int($j) < $fa Then $fa = Int($j)
	Next
	Return UBound($a)
EndFunc

; vom Klick in eine Zeile des markierten Patienten hinauf in seine erste (die mit dem Nachnamen)
Func SmartPixZurErstenZeile()
	Sleep(200)
	For $i = 1 To 500
		If SmartPixZelle() <> "" Then Return True
		Send("{UP}")
		Sleep(30)
	Next
	Return False
EndFunc

; Patient $ab (Nachname, Vorname, Geburtsdatum, Patienten-ID) in der Liste suchen: Suche leeren, oben anfangen, die
; Anfangsbuchstaben tippen (springt zum ersten so beginnenden Nachnamen) und von dort die Zeilen vergleichen; True
; mit dem Zellcursor in seiner ersten Zeile, @extended = 1, wenn ihm die gemerkte Ergaenzung fehlt
Func SmartPixZellenFinden($hListe, ByRef $ab, ByRef $sp)
	Local $anfang = StringRegExpReplace($ab[0], "\s.*$", ""), $r[5], $vorig = "", $gleich = 0, $passt, $ohne
	If $anfang = "" Then Return False
	HinweisStehend("Smart Pix-Liste: suche " & $ab[1] & " " & $ab[0] & " ...")
	SmartPixLog("suche " & SmartPixKurz($ab))
	If Not SmartPixListeOben($hListe) Then Return False
	Send($anfang, 1)
	Sleep(1500)
	For $i = 1 To 5000
		If _IsPressed("1B") Then Return False
		If Not SmartPixReiheLesen($r, $sp) Then
			SmartPixLog("suche: Zelle nicht lesbar")
			Return False
		EndIf
		If $r[0] <> "" Then
			If StringLeft(StringLower($r[0]), StringLen($anfang)) <> StringLower($anfang) Then ExitLoop
			$passt = SmartPixPasst($r, $ab)
			$ohne = @extended
			If $passt Then
				ToolTip("")
				SmartPixLog($ohne ? "gefunden, aber ohne die eingetragenen Werte (nicht gespeichert): noch einmal" : "gefunden")
				Return SetExtended($ohne, True)
			EndIf
		EndIf
		If SmartPixReiheText($r) = $vorig Then
			$gleich += 1
			If $gleich >= 10 Then ExitLoop
		Else
			$gleich = 0
		EndIf
		$vorig = SmartPixReiheText($r)
		Send("{DOWN}")
		Sleep(30)
	Next
	ToolTip("")
	SmartPixLog("suche: nicht gefunden")
	Return False
EndFunc

; den markierten Patienten oeffnen, pruefen, dass der Dialog $f zeigt, und $geb bzw. $id eintragen; "" wenn
; gespeichert, "schon da", wenn der Dialog den fehlenden Wert schon zeigt (aus der Zelle nicht kopiert), sonst der
; Grund. Ein in der Liste leer gelesenes Geburtsdatum, das der Dialog zeigt, laesst ihn aus (der Vorschlag beruhte
; dann nur auf dem Namen)
Func SmartPixZellenSetzen($hListe, ByRef $f, ByRef $geb, ByRef $id)
	Local $f2[4], $hDlg, $a, $b
	$gSmartPixWarnung = ""
	$gMeldung = ""
	$hDlg = SmartPixDialog($hListe, $f2)
	If Not $hDlg Then
		SmartPixDialogWeg()
		Return "nicht zu oeffnen" & (($gSmartPixWarnung = "") ? "" : " (Smart Pix: " & $gSmartPixWarnung & ")")
	EndIf
	For $i = 0 To 3
		$a = StringStripWS($f[$i], 3)
		$b = StringStripWS($f2[$i], 3)
		If $a = $b Then ContinueLoop
		If $i = 2 And $a = "" Then
			SmartPixLog("Geburtsdatum im Dialog, in der Liste leer gelesen: ausgelassen")
			SmartPixDialogZu($hDlg)
			Return "Geburtsdatum in der Liste nicht gelesen"
		EndIf
		If $i = 3 And $a = "" Then
			SmartPixLog("Patienten-ID im Dialog, in der Liste leer gelesen")
			$f[3] = $b
			$id = ""
			ContinueLoop
		EndIf
		SmartPixLog("Dialog zeigt einen anderen Patienten: " & SmartPixKurz($f2))
		SmartPixDialogZu($hDlg)
		Return "Dialog zeigt einen anderen Patienten"
	Next
	If $geb = "" And $id = "" Then
		SmartPixDialogZu($hDlg)
		Return "schon da"
	EndIf
	If SmartPixSetzen($hDlg, $geb, $id, $f[1]) Then Return ""
	Return "nicht gespeichert" & (($gSmartPixWarnung = "") ? "" : " (Smart Pix: " & $gSmartPixWarnung & ")")
EndFunc

; ein (zu spaet erschienener) Dialog "Patientendaten bearbeiten" steht noch offen: verwerfen, bis 5 s darauf warten
Func SmartPixDialogWeg($ms = 5000)
	Local $t = TimerInit(), $d
	Do
		$d = SmartPixFenster("[TITLE:Patientendaten bearbeiten; CLASS:Qt5QWindowIcon]")
		If $d Then
			SmartPixLog("Dialog erschien zu spaet, verworfen")
			Return SmartPixDialogZu($d)
		EndIf
		Sleep(250)
	Until TimerDiff($t) > $ms
	Return False
EndFunc

; nach dem Dialog zurueck in die Tabelle, in die erste Zeile des naechsten Patienten ($r, schon gelesen): der Fokus
; liegt dann auf einem Knopf, mit Umschalt+Tab zurueck in die Tabelle (dort steht der Zellcursor noch beim eben
; geoeffneten Patienten $f, die Zahl der Tabs wird gemerkt) und Pfeil runter; sonst in die Markierung klicken, sonst
; suchen
Func SmartPixZellenZurueck($hListe, $zurueck, ByRef $f, ByRef $r, ByRef $sp)
	Local $t[5], $z
	SmartPixDialogWeg(0)
	WinActivate($hListe)
	If Not WinWaitActive($hListe, "", 3) Then Return False
	Sleep(300)
	For $k = 1 To 8
		Send("+{TAB}")
		Sleep(150)
		If $gSpTabs > 0 And $k < $gSpTabs Then ContinueLoop
		$z = SmartPixZelle()
		If @error Or $z <> StringStripWS($f[0], 3) Then ContinueLoop
		If $gSpTabs <> $k Then SmartPixLog("Tabelle nach " & $k & "x Umschalt+Tab")
		$gSpTabs = $k
		Send("{DOWN}")
		Sleep(200)
		If SmartPixReiheLesen($t, $sp) And $t[0] = $r[0] And $t[1] = $r[1] And $t[2] = $r[2] And $t[3] = $r[3] Then Return True
		SmartPixLog("nach Umschalt+Tab in einer anderen Zeile: " & SmartPixKurz($t))
		ExitLoop
	Next
	If SmartPixWeiter($hListe, $zurueck) Then
		Sleep(200)
		If SmartPixReiheLesen($t, $sp) And $t[0] = $r[0] And $t[1] = $r[1] And $t[2] = $r[2] And $t[3] = $r[3] Then Return True
		SmartPixLog("nach dem Dialog in einer anderen Zeile: " & SmartPixKurz($t))
	EndIf
	Return SmartPixZellenFinden($hListe, $r, $sp)
EndFunc

; zum naechsten Patienten der Liste: ohne [SmartPixListe] ZurTabelle= wird die dunkle Unterkante der Markierung
; (alle Zeilen des markierten Patienten) in der Spalte Nachname gesucht, knapp darueber in dessen letzte Zeile
; geklickt (damit hat die Tabelle den Fokus) und Pfeil runter gedrueckt; sonst erst die Tasten aus ZurTabelle=
; ganze Liste: Suche leeren, eine Zeile anklicken (damit die Tabelle den Fokus hat) und mit Strg+Pos1 nach oben
Func SmartPixListeOben($hListe)
	If Not SmartPixKlick($hListe, Einst("SmartPix", "Suche", "357,-76"), 1) Then Return False
	Send("^a{DEL}")
	Sleep(Int(Einst("SmartPixListe", "ListeWarten", "5000")))
	If Not SmartPixKlick($hListe, Einst("SmartPix", "Zeile", "60,88"), 1) Then Return False
	Sleep(300)
	Send("^{HOME}")
	Sleep(500)
	Return True
EndFunc

; Smart Pix beenden (nur das Programm selbst, nicht das Connection Framework) und mit der Patientenliste neu starten;
; SmartPixListeOeffnen raeumt vor dem Start die Sperren dieses PCs auf. Handle der Liste oder 0
Func SmartPixNeustart()
	SmartPixLog("Neustart von Smart Pix")
	HinweisStehend("Smart Pix wird neu gestartet (Sperren in palock.xml) ...")
	Local $pl = ProcessList(SmartPixExe())
	For $i = 1 To $pl[0][0]
		ProcessClose($pl[$i][1])
		ProcessWaitClose($pl[$i][1], 10)
	Next
	Sleep(1000)
	Return SmartPixListeOeffnen()
EndFunc

; Sperrdatei von Smart Pix: palock.xml neben palist.xml ([SmartPixListe] Palist=)
Func SmartPixSperrDatei()
	Local $p = Einst("SmartPixListe", "Palist", "P:\Datenbanken\AccuChek\Diaries\palist.xml")
	Return StringLeft($p, StringInStr($p, "\", 0, -1)) & "palock.xml"
EndFunc

; so steht dieser PC in palock.xml (User="PC/Benutzer")
Func SmartPixIch()
	Return @ComputerName & "/" & @UserName
EndFunc

; Sperren laut palock.xml: Rueckgabe die anderer PCs als Text ("" = keine), @extended = LockCount der eigenen
; Listensperre (0 = keine)
Func SmartPixSperren()
	Local $t = FileRead(SmartPixSperrDatei()), $eigen = 0, $fremd = "", $u, $c
	Local $m = StringRegExp($t, '<PATIENT(?:LIST)?LOCK\b[^>]*>', 3)
	If @error Then Return SetExtended(0, "")
	For $e In $m
		$u = StringRegExp($e, 'User="([^"]*)"', 1)
		If @error Then ContinueLoop
		$c = StringRegExp($e, 'LockCount="(\d+)"', 1)
		If StringLower($u[0]) = StringLower(SmartPixIch()) Then
			If StringInStr($e, "PATIENTLISTLOCK") And IsArray($c) Then $eigen = Int($c[0])
		Else
			$fremd &= ", " & $u[0] & (StringInStr($e, "PATIENTLISTLOCK") ? " (Liste)" : " (Patient)")
		EndIf
	Next
	Return SetExtended($eigen, StringTrimLeft($fremd, 2))
EndFunc

; entfernt die Sperren dieses PCs aus palock.xml, aber nur, wenn Smart Pix hier nicht laeuft: hart beendet gibt es
; sie nicht frei und zaehlt sie beim naechsten Start weiter (LockCount 2, 3 ...), dann scheitert jedes Speichern
; ("Kein Schreibzugriff"); Sperren anderer PCs bleiben
Func SmartPixSperrenAufraeumen()
	If ProcessExists(SmartPixExe()) Then Return False
	Local $datei = SmartPixSperrDatei(), $t = FileRead($datei)
	If @error Then Return False
	Local $ich = StringRegExpReplace(SmartPixIch(), "([\\^$.|?*+()\[\]{}])", "\\$1")
	Local $neu = StringRegExpReplace($t, '(?i)[ \t]*<PATIENT(?:LIST)?LOCK\b[^>]*\bUser="' & $ich & '"[^>]*/>\r?\n?', "")
	If $neu == $t Then Return True
	$neu = StringRegExpReplace($neu, '<PATIENTLOCKS>\s*</PATIENTLOCKS>', '<PATIENTLOCKS/>')
	; 2 = ueberschreiben, 256 = UTF-8 ohne BOM wie bei Smart Pix
	Local $h = FileOpen($datei, 2 + 256)
	If $h = -1 Then Return False
	FileWrite($h, $neu)
	FileClose($h)
	Return True
EndFunc

; nach Strg+Pos1 zum Patienten $ab (Nachname, Vorname, Geburtsdatum, Patienten-ID wie gelesen) springen und auf den
; naechsten stellen: das Tippen des Nachnamens (bis zum ersten Leerzeichen, das waehlt sonst die Zeile) springt in der
; Tabelle zum ersten so beginnenden; von dort mit SmartPixWeiter weiter, bis alle vier Felder passen, solange der
; Nachname noch so beginnt. False, wenn nicht gefunden
Func SmartPixAnspringen($hListe, ByRef $ab, $zurueck)
	Local $anfang = StringRegExpReplace($ab[0], "\s.*$", ""), $f[4], $h, $passt, $ohne
	$gSpNochmal = False
	If $anfang = "" Then Return False
	HinweisStehend("Smart Pix-Liste: suche " & $ab[1] & " " & $ab[0] & " ...")
	SmartPixLog("suche " & SmartPixKurz($ab))
	Send($anfang, 1)
	Sleep(1500)
	For $i = 1 To 500
		If _IsPressed("1B") Then Return False
		; "Kein Schreibzugriff" beim Oeffnen ist meist voruebergehend: bis zu dreimal versuchen
		For $j = 1 To 3
			$h = SmartPixDialog($hListe, $f)
			If $h Then ExitLoop
			Sleep(2000 * $j)
			SmartPixInfoWeg()
			SmartPixWarnungWeg()
		Next
		If Not $h Then
			SmartPixLog("suche: Oeffnen gescheitert")
			Return False
		EndIf
		SmartPixDialogZu($h)
		$passt = SmartPixPasst($f, $ab)
		$ohne = @extended
		If $passt Then
			ToolTip("")
			; die gemerkte Ergaenzung ist in Smart Pix nicht angekommen: auf ihm stehenbleiben, er kommt noch einmal dran
			If $ohne Then
				SmartPixLog("gefunden, aber ohne die eingetragenen Werte (nicht gespeichert): noch einmal")
				$gSpNochmal = True
				Return True
			EndIf
			SmartPixLog("gefunden")
			Return SmartPixWeiter($hListe, $zurueck)
		EndIf
		; das Springen hat nicht geklappt bzw. schon an allen so beginnenden vorbei
		If StringLeft(StringLower($f[0]), StringLen($anfang)) <> StringLower($anfang) Then
			SmartPixLog("suche: an allen so beginnenden vorbei, nicht gefunden")
			Return False
		EndIf
		If Not SmartPixWeiter($hListe, $zurueck) Then Return False
		Sleep(Int(Einst("SmartPixListe", "Pause", "200")))
	Next
	Return False
EndFunc

; nach dem Eintragen gilt der Patient mit den neuen Werten (Wiederfinden nach einem Neustart, Stand, gleiche Daten
; bei weiteren Geraeten); $vs, $vid: eingetragenes Geburtsdatum bzw. Patienten-ID ("" = unveraendert)
Func SmartPixGemerkt(ByRef $f, ByRef $vorher, $vs, $vid)
	If $vs <> "" Then $f[2] = $vs
	If $vid <> "" Then $f[3] = $vid
	$vorher = $f[0] & "|" & $f[1] & "|" & $f[2] & "|" & $f[3]
EndFunc

; gelesener Patient $f = gemerkter $ab? Leerzeichen am Rand zaehlen nicht (beim Speichern womoeglich an den Vornamen
; angehaengt), ein gemerkt leeres Geburtsdatum bzw. leere Patienten-ID passt auf jeden Wert (inzwischen ergaenzt);
; ist es in Smart Pix leer, gemerkt aber nicht (Ergaenzung nicht gespeichert), passt es auch, mit @extended = 1
Func SmartPixPasst(ByRef $f, ByRef $ab)
	Local $a, $b, $verloren = 0
	For $i = 0 To 3
		$a = StringStripWS($ab[$i], 3)
		$b = StringStripWS($f[$i], 3)
		If $a = $b Or ($i >= 2 And $a = "") Then ContinueLoop
		If $i >= 2 And $b = "" Then
			$verloren = 1
			ContinueLoop
		EndIf
		Return False
	Next
	Return SetExtended($verloren, True)
EndFunc

; nach dem Wiederfinden noch $n Patienten weiter (die zuvor uebersprungenen); False, wenn es nicht weitergeht
Func SmartPixUeberspringen($hListe, $zurueck, $n)
	For $i = 1 To $n
		Sleep(Int(Einst("SmartPixListe", "Pause", "200")))
		If Not SmartPixWeiter($hListe, $zurueck) Then Return False
	Next
	Return True
EndFunc

Func SmartPixWeiter($hListe, $zurueck, $runter = True)
	If $zurueck <> "" Then
		Send($zurueck)
		Sleep(200)
		If $runter Then Send("{DOWN}")
		Return True
	EndIf
	Local $z = SmartPixPunkt($hListe, Einst("SmartPix", "Zeile", "60,88"))
	If @error Then Return False
	Local $unten = SmartPixPunkt($hListe, "0,-125"), $x = $z[0] + Int(Einst("SmartPixListe", "KanteX", "40"))
	Local $farbe = Int(Einst("SmartPixListe", "Kante", "0x0DA69E")), $p, $ok = False
	Opt("PixelCoordMode", 1)
	; von unten nach oben, damit die untere Kante gefunden wird; Qt zeichnet nach dem Dialog womoeglich erst spaeter
	For $i = 1 To 5
		Sleep(200)
		$p = PixelSearch($x, $unten[1], $x, $z[1] - 14, $farbe, 0x20)
		If Not @error Then
			$ok = True
			ExitLoop
		EndIf
	Next
	If Not $ok Then
		SmartPixKanteDiagnose($x, $z[1] - 14, $unten[1])
		Return False
	EndIf
	; folgt unter der gefundenen Kante noch Markierung (tuerkis getoent), ist es die Oberkante eines Patienten, der
	; unten aus dem Bild ragt: dann in die unterste sichtbare Zeile klicken
	Local $y = $p[1] - 8, $c
	If $p[1] + 4 < $unten[1] Then
		$c = PixelGetColor($x, $p[1] + 4)
		If BitAND(BitShift($c, 8), 0xFF) - BitShift($c, 16) > 8 Then $y = $unten[1] - 6
	EndIf
	Local $alt = MouseGetPos()
	Opt("MouseCoordMode", 1)
	MouseClick("left", $x, $y, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	Sleep(200)
	If $runter Then Send("{DOWN}")
	Return True
EndFunc

; Kante nicht gefunden: Farben entlang der Spalte (bei Wechsel) und ein Bild der Liste nach
; %APPDATA%\MOReiter\SmartPixKante.txt bzw. .png, zum Einstellen von Kante= und KanteX=
Func SmartPixKanteDiagnose($x, $oben, $unten)
	Local $t = "Spalte x=" & $x & ", y " & $oben & " bis " & $unten & @CRLF, $vorher = -1, $c
	For $y = $oben To $unten
		$c = PixelGetColor($x, $y)
		If $c <> $vorher Then $t &= "y=" & $y & " 0x" & Hex($c, 6) & @CRLF
		$vorher = $c
	Next
	DirCreate(@AppDataDir & "\MOReiter")
	Local $h = FileOpen(@AppDataDir & "\MOReiter\SmartPixKante.txt", 2)
	FileWrite($h, $t)
	FileClose($h)
	_ScreenCapture_Capture(@AppDataDir & "\MOReiter\SmartPixKante.png", $x - 300, $oben - 40, $x + 900, $unten, False)
EndFunc

; Fehler beim Durchlauf der ganzen Liste als Fenster (ein Hinweis an der Maus wird leicht uebersehen)
Func SmartPixFehler($t)
	ToolTip("")
	MsgBox(48 + 262144, "MOReiter: Smart Pix-Patientenliste", $t)
	Return False
EndFunc

; Patientennummern durch Leerzeichen getrennt (ohne Komma, wegen der CSV-Zelle)
Func SmartPixListeText(ByRef $a)
	Local $s = ""
	For $i = 0 To UBound($a) - 1
		$s &= ($s = "" ? "" : " ") & $a[$i]
	Next
	Return $s
EndFunc

Func SmartPixZeile($h, $nr, ByRef $f, $la, $mo, $qu, $vs, $vid, $ein)
	Local $z = $nr
	For $i = 0 To 3
		$z &= ";" & StringReplace($f[$i], ";", ",")
	Next
	FileWriteLine($h, $z & ";" & $la & ";" & $mo & ";" & $qu & ";" & $vs & ";" & $vid & ";" & $ein)
EndFunc

; Vorschlaege mit Eintragen=j aus dem Probelauf: Schluessel Nr|Nachname|Vorname (wie SmartPixNorm) ->
; Geburtsdatum|Patienten-ID (je leer, wenn kein Vorschlag)
Func SmartPixVorschlaege($datei, ByRef $vorschlag)
	Local $z = FileReadToArray($datei), $s
	If @error Then Return SmartPixFehler("Smart Pix-Liste: erst den Probelauf machen (" & $datei & " fehlt)")
	For $i = 1 To UBound($z) - 1
		$s = StringSplit($z[$i], ";", 2)
		If UBound($s) < 11 Then ContinueLoop
		If StringLower(StringStripWS($s[10], 3)) <> "j" Then ContinueLoop
		If Not StringRegExp($s[8], "^(\d\d\.\d\d\.\d{4})?$") Or Not StringRegExp($s[9], "^\d*$") Or $s[8] & $s[9] = "" Then ContinueLoop
		$vorschlag.Item($s[0] & "|" & SmartPixNorm($s[1]) & "|" & SmartPixNorm($s[2])) = $s[8] & "|" & $s[9]
	Next
	If $vorschlag.Count = 0 Then Return SmartPixFehler("Smart Pix-Liste: keine Zeile mit Eintragen=j in " & $datei & " (Probelauf von vor dieser Version?)")
	Return True
EndFunc

; liest die Patienten von Medical Office ("medoff") bzw. quelle in das gemeinsame Dictionary $d (die
; Patientennummer ist in beiden dieselbe): N|nach|vor, E|nach|erster Vorname, G|nach|geb -> Patientennummern
; (durch Komma getrennt), P|nr -> Geburtsdatum, F|nr -> Beginn des ersten Behandlungsfalls (JJJJMMTT, aus
; beiden der fruehere), M|nr bzw. Q|nr -> 1 (vorhanden), Z|nr -> 1 (als verstorben mit "zzz" gefuehrt);
; bei Fehler @error und der Fehlertext
Func DBPatienten($welche, $d)
	Local $cn = DBVerbinden($welche)
	If @error Then Return SetError(1, 0, $cn)
	Local $sql, $kz
	If $welche = "medoff" Then
		$kz = "M"
		; ohne die kuenstlichen Faelle vom 1.1.2025 (FVon 49308 = Tage seit 1.1.1890, Scheintyp 0, Tarif 0)
		$sql = "SELECT p.FSurogat, p.FNachname, p.FVorname, DATE_FORMAT(p.FGeburtsdatum,'%d.%m.%Y'), " & _
				"DATE_FORMAT(DATE_ADD('1890-01-01', INTERVAL f.v DAY),'%Y%m%d') FROM patstamm p LEFT JOIN " & _
				"(SELECT FPatnr, MIN(FVon) v FROM patfall WHERE NOT (FVon=49308 AND FScheintyp=0 AND FTarif=0) " & _
				"GROUP BY FPatnr) f ON f.FPatnr=p.FSurogat"
	Else
		$kz = "Q"
		; bhfb (Beginn des Behandlungsfalls) ist bei allen Faellen befuellt, auch bei denen aus Medical Office
		$sql = "SELECT n.pat_id, n.Nachname, n.Vorname, DATE_FORMAT(n.GebDat,'%d.%m.%Y'), DATE_FORMAT(f.b,'%Y%m%d') " & _
				"FROM namen n LEFT JOIN (SELECT pat_id, MIN(bhfb) b FROM faelle GROUP BY pat_id) f ON f.pat_id=n.pat_id"
	EndIf
	$gComFehler = ""
	$cn.CommandTimeout = 300
	Local $rs = $cn.Execute($sql)
	If $gComFehler <> "" Or Not IsObj($rs) Then
		Local $t = $gComFehler
		$cn.Close
		Return SetError(1, 0, "Abfrage der Patienten in " & $welche & " fehlgeschlagen: " & $t)
	EndIf
	Local $nr, $nach, $vor, $geb, $fb
	While Not $rs.EOF
		$nr = String($rs.Fields(0).Value)
		$nach = SmartPixNorm(String($rs.Fields(1).Value))
		$vor = SmartPixNorm(String($rs.Fields(2).Value))
		$geb = String($rs.Fields(3).Value)
		$fb = String($rs.Fields(4).Value)
		SmartPixDazu($d, "N|" & $nach & "|" & $vor, $nr)
		SmartPixDazu($d, "E|" & $nach & "|" & StringRegExpReplace($vor, " .*", ""), $nr)
		If $geb <> "" Then SmartPixDazu($d, "G|" & $nach & "|" & $geb, $nr)
		If Not $d.Exists("P|" & $nr) Or $d.Item("P|" & $nr) = "" Then $d.Item("P|" & $nr) = $geb
		If StringRegExp($fb, "^\d{8}$") Then
			If Not $d.Exists("F|" & $nr) Then
				$d.Item("F|" & $nr) = $fb
			ElseIf Int($fb) < Int($d.Item("F|" & $nr)) Then
				$d.Item("F|" & $nr) = $fb
			EndIf
		EndIf
		$d.Item($kz & "|" & $nr) = 1
		If StringRegExp(String($rs.Fields(1).Value), "(?i)^[*^\s]*zzz") Then $d.Item("Z|" & $nr) = 1
		$rs.MoveNext
	WEnd
	$rs.Close
	$cn.Close
	Return True
EndFunc

; haengt $w an die Komma-Liste unter $k an, wenn noch nicht enthalten
Func SmartPixDazu($d, $k, $w)
	If Not $d.Exists($k) Then
		$d.Item($k) = $w
	ElseIf $w <> "" And Not StringInStr("," & $d.Item($k) & ",", "," & $w & ",") Then
		$d.Item($k) = ($d.Item($k) = "" ? "" : $d.Item($k) & ",") & $w
	EndIf
EndFunc

; Patientennummern unter $k, nur die mit Geburtsdatum $geb (leer = alle) und vorhanden in $kz ("M", "Q", leer = beide)
Func SmartPixIds($d, $k, $geb = "", $kz = "")
	Local $erg[0], $l
	If Not $d.Exists($k) Then Return $erg
	$l = StringSplit($d.Item($k), ",", 2)
	For $i = 0 To UBound($l) - 1
		If $l[$i] = "" Then ContinueLoop
		If $geb <> "" And $d.Item("P|" & $l[$i]) <> $geb Then ContinueLoop
		If $kz <> "" And Not $d.Exists($kz & "|" & $l[$i]) Then ContinueLoop
		ReDim $erg[UBound($erg) + 1]
		$erg[UBound($erg) - 1] = $l[$i]
	Next
	Return $erg
EndFunc

; Spalte Medical Office ($kz "M") bzw. quelle ("Q"): "ja" (Name und, falls in Smart Pix vorhanden, Geburtsdatum),
; "ja, verstorben", "ja, erster Vorname", "ja, Vorname anders" (Nachname und Geburtsdatum), "Name, Geb. ..."
; (gleicher Name, anderes Geburtsdatum: wohl jemand anderes) oder ""
Func SmartPixAbgleich($d, ByRef $f, $kz)
	Local $nach = SmartPixNorm($f[0]), $vor = SmartPixNorm($f[1]), $erst = StringRegExpReplace($vor, " .*", ""), $ids
	$ids = SmartPixIds($d, "N|" & $nach & "|" & $vor, $f[2], $kz)
	If UBound($ids) > 0 Then Return "ja" & ($d.Exists("Z|" & $ids[0]) ? ", verstorben" : "")
	$ids = SmartPixIds($d, "E|" & $nach & "|" & $erst, $f[2], $kz)
	If UBound($ids) > 0 Then Return "ja, erster Vorname"
	If $f[2] = "" Then Return ""
	$ids = SmartPixIds($d, "G|" & $nach & "|" & $f[2], "", $kz)
	If UBound($ids) > 0 Then Return "ja, Vorname anders"
	$ids = SmartPixIds($d, "N|" & $nach & "|" & $vor, "", $kz)
	If UBound($ids) = 0 Then Return ""
	Local $s = "Name, Geb."
	For $i = 0 To UBound($ids) - 1
		$s &= " " & $d.Item("P|" & $ids[$i])
	Next
	Return $s
EndFunc

; Patienten aus Medical Office und quelle, die zu diesem Smart Pix-Eintrag passen: gleicher Name (sonst gleicher
; erster Vorname) und, falls in Smart Pix vorhanden, gleiches Geburtsdatum; ohne die, deren erster
; Behandlungsfall erst nach dem letzten Auslesen $bis (JJJJMMTT, leer = ohne diese Pruefung) begann, und ohne die,
; deren erster Fall mehr als [SmartPixListe] Vorlauf= Monate (Vorgabe 3, 0 = ohne) nach dem ersten Auslesen begann;
; das erste Auslesen steht nicht in palist.xml, $frueh (JJJJMMTT, 0 = unbekannt) ist das frueheste letzte Auslesen
; der Geraete, das erste lag nicht spaeter: so wird eher zu wenig als zu viel gestrichen
Func SmartPixKandidaten($d, ByRef $f, $bis, $frueh = 0)
	Local $nach = SmartPixNorm($f[0]), $vor = SmartPixNorm($f[1]), $erg[0], $fall
	Local $monate = Int(Einst("SmartPixListe", "Vorlauf", "3")), $grenze = 0
	If $monate > 0 And Int($frueh) > 0 Then
		; JJJJMMTT plus $monate Monate (Tag bleibt, fuer den Vergleich genuegt das)
		Local $mm = Int(StringMid($frueh, 5, 2)) - 1 + $monate
		$grenze = (Int(StringLeft($frueh, 4)) + Floor($mm / 12)) * 10000 + (Mod($mm, 12) + 1) * 100 + Int(StringRight($frueh, 2))
	EndIf
	Local $ids = SmartPixIds($d, "N|" & $nach & "|" & $vor, $f[2])
	If UBound($ids) = 0 Then $ids = SmartPixIds($d, "E|" & $nach & "|" & StringRegExpReplace($vor, " .*", ""), $f[2])
	For $i = 0 To UBound($ids) - 1
		If $d.Exists("F|" & $ids[$i]) Then
			$fall = Int($d.Item("F|" & $ids[$i]))
			If $bis <> "" And $fall > Int($bis) Then ContinueLoop
			If $grenze > 0 And $fall > $grenze Then ContinueLoop
		EndIf
		ReDim $erg[UBound($erg) + 1]
		$erg[UBound($erg) - 1] = $ids[$i]
	Next
	Return $erg
EndFunc

; letztes Auslesen je Tagebuch aus palist.xml von Smart Pix (dort unverschluesselt): "{...}.dia" (klein) -> JJJJMMTT
Func SmartPixAuslesen($datei)
	Local $d = ObjCreate("Scripting.Dictionary")
	Local $t = FileRead($datei)
	If @error Then Return SetError(1, 0, $d)
	Local $teile = StringSplit($t, "<PATIENT ", 1), $dia, $daten, $max, $min, $z
	For $i = 2 To $teile[0]
		$dia = StringRegExp($teile[$i], 'DiaryFileName="([^"]+)"', 1)
		If @error Then ContinueLoop
		$daten = StringRegExp($teile[$i], 'LastUpload="(\d{4})-(\d\d)-(\d\d)"', 3)
		If @error Then ContinueLoop
		$max = 0
		$min = 99999999
		For $j = 0 To UBound($daten) - 3 Step 3
			$z = Int($daten[$j] & $daten[$j + 1] & $daten[$j + 2])
			If $z > $max Then $max = $z
			If $z > 0 And $z < $min Then $min = $z
		Next
		$d.Item(StringLower($dia[0])) = String($max)
		; frueheste letzte Auslesen der Geraete: das erste Auslesen lag nicht spaeter (SmartPixKandidaten)
		If $min < 99999999 Then $d.Item("E|" & StringLower($dia[0])) = String($min)
		; auch je interner Nummer, fuer die Zuordnung ueber die Protokolle (SmartPixProtokolle)
		$dia = StringRegExp($teile[$i], '^UID="(\d+)"', 1)
		If Not @error Then
			$d.Item("U|" & $dia[0]) = String($max)
			If $min < 99999999 Then $d.Item("V|" & $dia[0]) = String($min)
		EndIf
	Next
	Return $d
EndFunc

; Zuordnung Patient -> interne Nummer (UID in palist.xml) aus den Protokollen von Smart Pix neben palist.xml
; (Log_JJMM.AccessInterface.csv: bei jedem Oeffnen und Auslesen "-PUID 'n' -lname1 '..' -fname2 '..' -bdate '..'
; -extid '..'"), nach $d: "I|Patienten-ID", "N|Nachname|Vorname|TT.MM.JJJJ" und "M|Nachname|Vorname" (Namen wie
; SmartPixNorm) -> UID, "" bei mehreren; "?" -> JJJJMMTT des Beginns der Protokolle
Func SmartPixProtokolle($palist, $d)
	Local $verz = StringLeft($palist, StringInStr($palist, "\", 0, -1)), $such = FileFindFirstFile($verz & "Log_????.AccessInterface.csv")
	If $such = -1 Then Return
	Local $f, $t, $m, $erst = "9999", $k, $g, $z
	While 1
		$f = FileFindNextFile($such)
		If @error Then ExitLoop
		If StringMid($f, 5, 4) < $erst Then $erst = StringMid($f, 5, 4)
		$t = FileRead($verz & $f)
		$m = StringRegExp($t, "-PUID '(\d+)' -lname1 '([^']*)' -fname2 '([^']*)' -bdate '([^']*)'((?: -extid '[^']*')?)", 3)
		If @error Then ContinueLoop
		For $i = 0 To UBound($m) - 5 Step 5
			; Geburtsdatum als JJJJ-MM-TT oder TT.MM.JJJJ, sonst leer
			$z = StringRegExp($m[$i + 3], "^(\d{4})-(\d\d)-(\d\d)$", 1)
			If @error Then
				$g = StringRegExp($m[$i + 3], "^\d\d\.\d\d\.\d{4}$") ? $m[$i + 3] : ""
			Else
				$g = $z[2] & "." & $z[1] & "." & $z[0]
			EndIf
			$k = SmartPixNorm($m[$i + 1]) & "|" & SmartPixNorm($m[$i + 2])
			SmartPixUidDazu($d, "N|" & $k & "|" & $g, $m[$i])
			SmartPixUidDazu($d, "M|" & $k, $m[$i])
			; die letzte Gruppe ist immer belegt (sonst fehlte sie im Ergebnis und verschoebe die folgenden)
			$z = StringStripWS(StringRegExpReplace($m[$i + 4], "^ -extid '|'$", ""), 3)
			If $z <> "" Then SmartPixUidDazu($d, "I|" & $z, $m[$i])
		Next
	WEnd
	FileClose($such)
	If $erst <> "9999" Then $d.Item("?") = "20" & $erst & "01"
EndFunc

; Schluessel -> UID; dieselbe UID mehrfach ist gleich, eine andere macht den Schluessel mehrdeutig ("")
Func SmartPixUidDazu($d, $k, $uid)
	If Not $d.Exists($k) Then
		$d.Item($k) = $uid
	ElseIf $d.Item($k) <> $uid Then
		$d.Item($k) = ""
	EndIf
EndFunc

; letztes Auslesen (JJJJMMTT) des Patienten mit den Feldern $f ueber die Protokolle: erst ueber die Patienten-ID, dann
; Name und Geburtsdatum, dann nur den Namen; kommt der Name in den Protokollen gar nicht vor, wurde er seit deren
; Beginn nicht ausgelesen, dann gilt dieser Beginn als spaetestes Datum. "" wenn unbekannt bzw. mehrdeutig;
; @extended: fruehestes letztes Auslesen der Geraete (JJJJMMTT, 0 = unbekannt)
Func SmartPixLetztesAuslesen($puids, $ausgelesen, ByRef $f)
	If $puids.Count = 0 Then Return SetExtended(0, "")
	Local $k = SmartPixNorm($f[0]) & "|" & SmartPixNorm($f[1]), $schl[3] = ["I|" & $f[3], "N|" & $k & "|" & $f[2], "M|" & $k], $uid
	For $i = 0 To 2
		If $i = 0 And $f[3] = "" Then ContinueLoop
		If Not $puids.Exists($schl[$i]) Then ContinueLoop
		$uid = $puids.Item($schl[$i])
		If $uid = "" Then Return SetExtended(0, "")
		If $ausgelesen.Exists("U|" & $uid) Then Return SetExtended($ausgelesen.Exists("V|" & $uid) ? _
				Int($ausgelesen.Item("V|" & $uid)) : 0, $ausgelesen.Item("U|" & $uid))
		Return SetExtended(0, "")
	Next
	Return SetExtended(0, $puids.Exists("?") ? $puids.Item("?") : "")
EndFunc

; Texte der Markierungen (Tabellen patmark/markier) des Patienten in Medical Office, sortiert nach Nummer der
; Markierung; Zugangsdaten [medoff] uid=, pwd= aus der Zugang.ini von Dateilese ([Allgemein] Zugang=), Server
; wie bei Dateilese aus der Registry (MOServer, sonst MORes); bei Fehler @error und der Fehlertext
Func MOMarkierungen($nach, $vor, $geb)
	Local $d = StringSplit($geb, ".", 2)
	If UBound($d) <> 3 Then Return SetError(1, 0, "Geburtsdatum " & $geb & " unlesbar")
	Local $sql = "SELECT DISTINCT m.FText, m.FSurogat FROM patmark pm JOIN markier m ON pm.FMarkiernr=m.FSurogat" & _
			" JOIN patstamm p ON p.FSurogat=pm.FPatnr WHERE p.FNachname='" & SqlText($nach) & "' AND p.FVorname='" & _
			SqlText($vor) & "' AND DATE(p.FGeburtsdatum)=DATE(" & $d[2] & $d[1] & $d[0] & ") ORDER BY m.FSurogat"
	Local $cn = DBVerbinden("medoff")
	If @error Then Return SetError(1, 0, $cn)
	$gComFehler = ""
	Local $rs = $cn.Execute($sql), $erg[0]
	If $gComFehler <> "" Or Not IsObj($rs) Then
		Local $t = $gComFehler
		$cn.Close
		Return SetError(1, 0, "Abfrage der Markierungen fehlgeschlagen: " & $t)
	EndIf
	While Not $rs.EOF
		ReDim $erg[UBound($erg) + 1]
		$erg[UBound($erg) - 1] = StringStripWS($rs.Fields(0).Value, 3)
		$rs.MoveNext
	WEnd
	$rs.Close
	$cn.Close
	Return $erg
EndFunc

; Patientennummer (patstamm.FSurogat) in Medical Office zu Name und Geburtsdatum; bei Fehler oder nicht genau
; einem Treffer @error und der Grund
Func MOPatNr($nach, $vor, $geb)
	Local $d = StringSplit($geb, ".", 2)
	If UBound($d) <> 3 Then Return SetError(1, 0, "Geburtsdatum " & $geb & " unlesbar")
	Local $cn = DBVerbinden("medoff")
	If @error Then Return SetError(1, 0, $cn)
	$gComFehler = ""
	Local $rs = $cn.Execute("SELECT FSurogat FROM patstamm WHERE FNachname='" & SqlText($nach) & "' AND FVorname='" & _
			SqlText($vor) & "' AND DATE(FGeburtsdatum)=DATE(" & $d[2] & $d[1] & $d[0] & ")"), $nr = "", $n = 0
	If $gComFehler <> "" Or Not IsObj($rs) Then
		Local $t = $gComFehler
		$cn.Close
		Return SetError(1, 0, "Abfrage fehlgeschlagen: " & $t)
	EndIf
	While Not $rs.EOF
		$nr = String($rs.Fields(0).Value)
		$n += 1
		$rs.MoveNext
	WEnd
	$rs.Close
	$cn.Close
	If $n <> 1 Then Return SetError(1, 0, $n & " Treffer")
	Return $nr
EndFunc

; ADODB-Verbindung zu "medoff" (Medical Office, Server wie bei Dateilese aus der Registry: MOServer, sonst MORes)
; oder "quelle" ([SmartPixListe] QuelleServer=, QuellePort=); Zugangsdaten [medoff] bzw. [quelle] uid=, pwd= aus
; der Zugang.ini von Dateilese ([Allgemein] Zugang=); bei Fehler @error und der Fehlertext
Func DBVerbinden($welche)
	Local $zugang = Einst("Allgemein", "Zugang", "C:\Program Files (x86)\Dateilesen\Zugang.ini")
	Local $uid = IniRead($zugang, $welche, "uid", ""), $pwd = IniRead($zugang, $welche, "pwd", "")
	If $uid = "" Then Return SetError(1, 0, "Keine Zugangsdaten zu " & $welche & " ([" & $welche & "] uid=, pwd= in " & $zugang & ")")
	Local $verb = "DRIVER={MySQL ODBC 8.0 Unicode Driver};option=0;database=" & $welche & ";uid=" & $uid & ";pwd=" & $pwd & ";port="
	Local $server[2]
	If $welche = "medoff" Then
		$verb &= Einst("Allgemein", "MOPort", "2020") & ";server="
		$server[0] = RegWertDL("MOServer", "wser")
		$server[1] = RegWertDL("MORes", "wres")
	Else
		$verb &= Einst("SmartPixListe", "QuellePort", "3306") & ";server="
		ReDim $server[1]
		$server[0] = Einst("SmartPixListe", "QuelleServer", "linux1")
	EndIf
	Local $cn = ObjCreate("ADODB.Connection"), $fehler = ""
	If Not IsObj($cn) Then Return SetError(1, 0, "ADODB nicht verfuegbar")
	For $srv In $server
		$gComFehler = ""
		$cn.ConnectionTimeout = 5
		$cn.Open($verb & $srv)
		If $gComFehler = "" And $cn.State = 1 Then Return $cn
		$fehler &= @CRLF & $srv & ": " & $gComFehler
	Next
	Return SetError(1, 0, "Keine Verbindung zur Datenbank " & $welche & ":" & $fehler)
EndFunc

; Wert aus HKCU\Software\GSProducts\Dateilese (Einstellungen von Dateilese), sonst $vorgabe
Func RegWertDL($wert, $vorgabe)
	Local $v = RegRead("HKCU\Software\GSProducts\Dateilese", $wert)
	If @error Or $v = "" Then Return $vorgabe
	Return $v
EndFunc

; Text fuer ein SQL-Literal in '...' (MySQL: Backslash und Hochkomma verdoppeln)
Func SqlText($t)
	Return StringReplace(StringReplace($t, "\", "\\"), "'", "''")
EndFunc

Func ComFehler($o)
	$gComFehler = StringStripWS($o.description, 3)
	If $gComFehler = "" Then $gComFehler = "COM-Fehler 0x" & Hex($o.number, 8)
EndFunc

; Stand des laufenden Programms: Aenderungszeit der exe (= Kompilierzeit, Kopieren erhaelt sie), TT.MM.JJJJ hh:mm
Func Stand()
	Local $t = FileGetTime(@AutoItExe, 0)
	If @error Then Return "?"
	Return $t[2] & "." & $t[1] & "." & $t[0] & " " & $t[3] & ":" & $t[4]
EndFunc

; Liste der eingerichteten Aufrufe fuer die Hilfe
Func AufrufeText()
	If UBound($gAufrufe) = 0 Then Return "(hier keine eingerichtet)"
	Local $t = ""
	For $i = 0 To UBound($gAufrufe) - 1
		$t &= StringFormat("%-20s", $gAufrufe[$i][0]) & $gAufrufe[$i][3] & @CRLF
	Next
	Return StringTrimRight($t, 2)
EndFunc

Func Einst($sektion, $schluessel, $def)
	Local $v = IniRead($Ini, $sektion, $schluessel, Chr(1))
	If $v <> Chr(1) Then Return $v
	Return IniRead($IniVorgaben, $sektion, $schluessel, $def)
EndFunc

; Strg+Alt+B: Laborbefunde des Patienten im Firefox, ebenso andere Portale (Abschnitt $s in der INI, z.B.
; [LibreView], aufgerufen ueber [Aufrufe] Strg+L=Portal:LibreView). MOReiter waehlt den Tab mit dem Portal
; (oder einen neuen) und oeffnet dort die Portaladresse mit dem Auftrag "#mor=Nachname;Vorname;Geburtsdatum;
; Kennung;Nr;PC-Name;Art;Tage"; die Arbeit auf der Seite macht das Benutzerskript MOReiter-<Abschnitt>.user.js
; (Tampermonkey), das seinen Fortschritt im Seitentitel "MOR:Nr:Zustand" meldet.
Func Labor()
	Return Portal("Labor")
EndFunc

Func Portal($s)
	; Zusatzfunktion fuer ein bestimmtes Portal, nur mit Eintrag in der INI
	Local $portal = Einst($s, "Url", "")
	If $portal = "" Then Return Meldung($s & " nicht eingerichtet ([" & $s & "] Url= in der INI, siehe Hilfe)")
	Local $hMO = WinGetHandle($MOFenster)
	If @error And Not $gLink Then Return MOStarten()
	If Not ModifierLos() Then Return Meldung("Strg und Alt bitte loslassen")
	Local $nach, $vor, $geb
	If Not PatientDaten($hMO, $nach, $vor, $geb) Then Return Meldung("Name und Geburtsdatum des Patienten nicht gefunden (Seite Personalien)")
	; beim ersten Mal auf dem PC: Tampermonkey und das Benutzerskript gleich einrichten, ohne erst
	; vergeblich auf eine Antwort zu warten; danach geht es mit dem Auftrag weiter
	If Not TampermonkeyDa() Then
		If Not AddOnEinrichten() Then Return False
		If Not SkriptEinrichten($s) Then Return False
	EndIf
	Local $r = PortalAuftrag($s, $portal, $nach, $vor, $geb)
	If Not @extended Then Return $r
	; keine Antwort, obwohl Tampermonkey da ist: das Skript fehlt (oder ist abgeschaltet)
	If Not TampermonkeyDa() And Not AddOnEinrichten() Then Return False
	If Not SkriptEinrichten($s) Then Return False
	$r = PortalAuftrag($s, $portal, $nach, $vor, $geb)
	If @extended Then Return Meldung("Keine Antwort vom Benutzerskript ""MOReiter " & $s & """: in Tampermonkey eingeschaltet?")
	Return $r
EndFunc

; Download-Ordner des Benutzers (Shell-Ordner "Downloads", sonst %USERPROFILE%\Downloads)
Func DownloadOrdner()
	Local $d = RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", _
			"{374DE290-123F-4565-9164-39C4925E467B}")
	If $d <> "" Then
		Local $e = DllCall("kernel32.dll", "dword", "ExpandEnvironmentStringsW", "wstr", $d, "wstr", "", "dword", 32767)
		If Not @error And $e[0] Then $d = $e[2]
	EndIf
	If $d = "" Or Not FileExists($d) Then $d = @UserProfileDir & "\Downloads"
	Return StringRegExpReplace($d, "\\$", "")
EndFunc

; Namen der PDFs im Download-Ordner als "|a.pdf|b.pdf|"
Func DownloadPdfs()
	Local $erg = "|", $such = FileFindFirstFile(DownloadOrdner() & "\*.pdf")
	If $such = -1 Then Return $erg
	While 1
		Local $f = FileFindNextFile($such)
		If @error Then ExitLoop
		If Not @extended Then $erg &= $f & "|"
	WEnd
	FileClose($such)
	Return $erg
EndFunc

; nach "fertig": ein danach im Download-Ordner neu erscheinendes PDF des Patienten nach [Portal] Ablage=
; verschieben; nur Dateien, deren Name zu [Portal] Datei= passt (regulaerer Ausdruck, ganzer Name, gross/klein
; egal; {vor} und {nach} stehen fuer Vor- und Nachname, s. NameMuster, {TT}, {MM}, {JJJJ} fuer das heutige
; Datum); ohne Datei= wird nichts verschoben
Func AblageMerken($s, $nach, $vor, $geb, $zeitraum = "")
	Local $ziel = Einst($s, "Ablage", ""), $muster = Einst($s, "Datei", "")
	If $ziel = "" Or $muster = "" Then Return
	Local $erster = StringRegExpReplace(StringStripWS($vor, 3), "\s.*$", "")
	$muster = StringReplace($muster, "{vor}", "(?:" & NameMuster($vor) & "|" & NameMuster($erster) & ")")
	$muster = StringReplace($muster, "{nach}", "(?:" & NameMuster($nach) & ")")
	; Herunterladedatum (heute)
	$muster = StringReplace(StringReplace(StringReplace($muster, "{TT}", @MDAY), "{MM}", @MON), "{JJJJ}", @YEAR)
	Local $n = UBound($gAblagen)
	ReDim $gAblagen[$n + 1][9]
	$gAblagen[$n][0] = $s
	$gAblagen[$n][1] = StringRegExpReplace($ziel, "\\$", "")
	$gAblagen[$n][2] = TimerInit()
	$gAblagen[$n][3] = DownloadPdfs()
	$gAblagen[$n][4] = ""
	$gAblagen[$n][5] = "(?i)^" & $muster & "$"
	$gAblagen[$n][6] = AblageNeuerName($s, $nach, $vor, $geb)
	$gAblagen[$n][7] = $zeitraum
	$gAblagen[$n][8] = $nach & @LF & $vor & @LF & $geb
EndFunc

; Name= : beim Verschieben so umbenennen, z.B. "{nach}, {vor}, {geb}, Glooko {TT}.{MM}.{JJJJ}.pdf";
; {geb} = Geburtsdatum TT.MM.JJJJ; Zeichen, die in Dateinamen nicht erlaubt sind, fallen weg; {von}, {bis}
; (Zeitraum des Berichts) bleiben stehen, s. NameMitZeitraum
Func AblageNeuerName($s, $nach, $vor, $geb)
	Local $name = Einst($s, "Name", "")
	If $name = "" Then Return ""
	Local $weg = "[\\/:*?""<>|]+"
	$name = StringReplace($name, "{vor}", StringStripWS(StringRegExpReplace($vor, $weg, ""), 7))
	$name = StringReplace($name, "{nach}", StringStripWS(StringRegExpReplace($nach, $weg, ""), 7))
	$name = StringReplace($name, "{geb}", $geb)
	Return StringReplace(StringReplace(StringReplace($name, "{TT}", @MDAY), "{MM}", @MON), "{JJJJ}", @YEAR)
EndFunc

; setzt den Zeitraum "TT.MM.JJJJ-TT.MM.JJJJ" fuer {von}, {bis} ein; ohne ihn entfaellt "{von}-{bis}"
Func NameMitZeitraum($name, $zeitraum)
	Local $zr = StringRegExp($zeitraum, "^(\d\d\.\d\d\.\d{4})-(\d\d\.\d\d\.\d{4})$", 1)
	If @error Then
		$name = StringRegExpReplace($name, "\s*\{von\}\s*-\s*\{bis\}", "")
		Return StringReplace(StringReplace($name, "{von}", ""), "{bis}", "")
	EndIf
	Return StringReplace(StringReplace($name, "{von}", $zr[0]), "{bis}", $zr[1])
EndFunc

; Zeitraum des Berichts von der ersten Seite des PDFs als "TT.MM.JJJJ-TT.MM.JJJJ", "" wenn nicht zu finden;
; braucht pdftotext.exe (Xpdf-Werkzeuge) unter [Allgemein] PdfText=. Erkannt werden z.B.
; "Do. 10. Sep. 2026 - Fr. 9. Okt. 2026" (Dexcom), "10 Jan 2024 - 6 Feb 2024" (Libre),
; "Fr, 19. Sep. 2025 - Do, 02. Okt. 2025" (Glooko) und "19.09.2025 - 02.10.2025"
Func PdfZeitraum($datei)
	Local $exe = Einst("Allgemein", "PdfText", "")
	If $exe = "" Or Not FileExists($exe) Then Return ""
	; 2 = $STDOUT_CHILD
	Local $pid = Run('"' & $exe & '" -l 1 -enc UTF-8 "' & $datei & '" -', "", @SW_HIDE, 2)
	If Not $pid Then Return ""
	Local $bin = Binary(""), $t = TimerInit()
	While TimerDiff($t) < 15000
		$bin &= StdoutRead($pid, False, True)
		If @error Then ExitLoop
		Sleep(50)
	WEnd
	If ProcessExists($pid) Then ProcessClose($pid)
	Local $txt = BinaryToString($bin, 4)
	; beide Daten in derselben Zeile (sonst naehme es z.B. bei Glooko das Geburtsdatum davor)
	Local $m = StringRegExp($txt, "(\d{1,2})\.(\d{1,2})\.(\d{4})[ \t]*[^\d\r\n]{1,3}[ \t]*(\d{1,2})\.(\d{1,2})\.(\d{4})", 1)
	If Not @error Then Return StringFormat("%02d.%02d.%04d-%02d.%02d.%04d", $m[0], $m[1], $m[2], $m[3], $m[4], $m[5])
	$m = StringRegExp($txt, "(\d{1,2})\.?\s+([^\s\d.,]{3,})\.?\s+(\d{4})[^\d\r\n]{1,15}?(\d{1,2})\.?\s+([^\s\d.,]{3,})\.?\s+(\d{4})", 3)
	For $i = 0 To UBound($m) - 6 Step 6
		Local $m1 = MonatNr($m[$i + 1]), $m2 = MonatNr($m[$i + 4])
		If $m1 And $m2 Then Return StringFormat("%02d.%02d.%04d-%02d.%02d.%04d", $m[$i], $m1, $m[$i + 2], $m[$i + 3], $m2, $m[$i + 5])
	Next
	Return ""
EndFunc

; Monat 1-12 aus einem deutschen oder englischen Namen bzw. seiner Abkuerzung, 0 = keiner
Func MonatNr($w)
	$w = StringLower(StringReplace($w, ChrW(228), "ae"))
	Local $l = StringSplit("jan|feb|mae|apr|mai|jun|jul|aug|sep|okt|nov|dez", "|", 2)
	Local $e = StringSplit("jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", "|", 2)
	For $i = 0 To 11
		If StringLeft($w, 3) = $l[$i] Or StringLeft($w, 3) = $e[$i] Then Return $i + 1
	Next
	Return 0
EndFunc

; regulaerer Ausdruck fuer einen Namen, wie ihn ein Portal in Dateinamen schreiben kann: ae/oe/ue/ss oder
; a/o/u statt Umlaut bzw. sz, Leerzeichen und Bindestrich auch weggelassen oder als Unterstrich
Func NameMuster($name)
	Local $erg = "", $z = StringSplit(StringStripWS($name, 3), "", 2)
	For $c In $z
		Switch $c
			Case ChrW(228), ChrW(196)
				$erg &= "(?:" & ChrW(228) & "|ae|a)"
			Case ChrW(246), ChrW(214)
				$erg &= "(?:" & ChrW(246) & "|oe|o)"
			Case ChrW(252), ChrW(220)
				$erg &= "(?:" & ChrW(252) & "|ue|u)"
			Case ChrW(223)
				$erg &= "(?:" & ChrW(223) & "|ss)"
			Case " ", "-"
				$erg &= "[ _-]?"
			Case Else
				$erg &= StringRegExp($c, "[A-Za-z0-9]") ? $c : "\" & $c
		EndSwitch
	Next
	Return $erg
EndFunc

; in der Hauptschleife alle 2 s: neues, fertig geschriebenes PDF (keine .part-Datei mehr, Groesse 3 s
; unveraendert) verschieben; je Auftrag eines, hoechstens 30 min lang
Func AblagenPruefen()
	$gAblagenGeprueft = TimerInit()
	Local $dl = DownloadOrdner(), $jetzt = DownloadPdfs()
	For $i = UBound($gAblagen) - 1 To 0 Step -1
		If TimerDiff($gAblagen[$i][2]) > 30 * 60000 Then
			AblageWeg($i)
			ContinueLoop
		EndIf
		For $f In StringSplit(StringTrimLeft(StringTrimRight($jetzt, 1), 1), "|", 2)
			If $f = "" Or StringInStr($gAblagen[$i][3], "|" & $f & "|") Then ContinueLoop
			If Not StringRegExp($f, $gAblagen[$i][5]) Then ContinueLoop
			If FileExists($dl & "\" & $f & ".part") Then ContinueLoop
			Local $gr = FileGetSize($dl & "\" & $f), $alt = StringSplit($gAblagen[$i][4], ";", 2)
			If UBound($alt) < 3 Or $alt[0] <> $f Or Number($alt[1]) <> $gr Then
				$gAblagen[$i][4] = $f & ";" & $gr & ";" & TimerInit()
				ContinueLoop
			EndIf
			If $gr = 0 Or TimerDiff(Number($alt[2])) < 3000 Then ContinueLoop
			Local $neu = $gAblagen[$i][6], $zeit = $gAblagen[$i][7]
			If $neu <> "" Then
				Local $zr = StringInStr($neu, "{von}") ? PdfZeitraum($dl & "\" & $f) : ""
				If $zr <> "" Then $zeit = $zr
				$neu = NameMitZeitraum($neu, $zeit)
			EndIf
			Local $ziel = AblageName($gAblagen[$i][1], ($neu = "") ? $f : $neu)
			If FileMove($dl & "\" & $f, $ziel) Then
				Local $pat = StringSplit($gAblagen[$i][8], @LF, 2)
				If UBound($pat) = 3 Then $ziel = DokEinordnen($gAblagen[$i][0], $ziel, $pat[0], $pat[1], $pat[2], $zeit)
				Hinweis($gAblagen[$i][0] & ": " & $f & " als " & $ziel & " abgelegt")
				; fuer andere Auftraege nicht mehr neu
				For $j = 0 To UBound($gAblagen) - 1
					$gAblagen[$j][3] &= $f & "|"
				Next
				AblageWeg($i)
			EndIf
			ExitLoop
		Next
	Next
EndFunc

; Zielname, bei schon vorhandener Datei mit " (2)", " (3)" ...
Func AblageName($verz, $f)
	Local $ziel = $verz & "\" & $f, $stamm = StringRegExpReplace($f, "(?i)\.pdf$", ""), $k = 2
	While FileExists($ziel)
		$ziel = $verz & "\" & $stamm & " (" & $k & ").pdf"
		$k += 1
	WEnd
	Return $ziel
EndFunc

Func AblageWeg($i)
	Local $n = UBound($gAblagen)
	For $j = $i To $n - 2
		For $k = 0 To UBound($gAblagen, 2) - 1
			$gAblagen[$j][$k] = $gAblagen[$j + 1][$k]
		Next
	Next
	ReDim $gAblagen[$n - 1][9]
EndFunc

; steht dieser PC in der Komma-Liste $liste (gross/klein egal)?
Func PCInListe($liste)
	For $pc In StringSplit($liste, ",", 2)
		If StringStripWS($pc, 3) <> "" And StringStripWS($pc, 3) = @ComputerName Then Return True
	Next
	Return False
EndFunc

; Voreinstellungen je Portal: Seitentitel, Adresse des Benutzerskripts
Func PortalTitel($s)
	Return Einst($s, "Titel", ($s = "Labor") ? $defLaborTitel : $s)
EndFunc

Func PortalSkript($s)
	Return Einst($s, "Skript", ($s = "Labor") ? $defLaborSkript : $defSkriptBasis & $s & ".user.js")
EndFunc

; schickt den Auftrag ans Portal und verfolgt ihn; @extended = 1: keine Antwort vom Benutzerskript
Func PortalAuftrag($s, $portal, $nach, $vor, $geb)
	Local $nr = Random(1000, 9999, 1)
	; Art je PC: auf den in SpeichernAuf= genannten PCs "speichern" (z.B. Dexcom: Bericht als PDF), sonst Art=
	Local $art = PCInListe(Einst($s, "SpeichernAuf", "")) ? "speichern" : Einst($s, "Art", "")
	Local $url = $portal & "#mor=" & UrlKodiert($nach) & ";" & UrlKodiert($vor) & ";" & UrlKodiert($geb) & ";" _
			& UrlKodiert(Einst($s, "Kennung", "")) & ";" & $nr & ";" & UrlKodiert(@ComputerName) & ";" _
			& UrlKodiert($art) & ";" & UrlKodiert(Einst($s, "Tage", ""))
	Local $hFF = FirefoxOeffnen($url, PortalTitel($s), $s)
	If Not $hFF Then Return False
	Local $t = TimerInit(), $max = Int(Einst($s, "Zeit", $defLaborZeit)), $z = "", $pw = False, $code = False
	; Zwischenstaende abwarten; bleibt jede Meldung aus, laeuft das Benutzerskript nicht
	; $gemeldet: das Skript hat sich mindestens einmal gemeldet; laeuft danach die Zeit ab, fehlt es nicht
	Local $tippen = 0, $gemeldet = False, $enter = 0, $liste = 0, $selbst = False, $zeitraum = ""
	While TimerDiff($t) < $max
		Local $warte = ($z = "") ? 15000 : $max - TimerDiff($t)
		; "passworttippen" bleibt stehen, wenn das Zeichen nicht im Feld ankam: nach 3 s noch einmal
		If $z = "passworttippen" Then $warte = 3000
		If $z = "enter" Then $warte = 6000
		If $z = "passwort" Then $warte = 3000
		Local $neu = FFZustand($hFF, $nr, $z, $warte)
		; "fertig|TT.MM.JJJJ-TT.MM.JJJJ": Zeitraum des Berichts fuer den Dateinamen (Name= {von}, {bis})
		Local $strich = StringInStr($neu, "|")
		If $strich Then
			$zeitraum = StringMid($neu, $strich + 1)
			$neu = StringLeft($neu, $strich - 1)
		EndIf
		If $neu = "" And ($z = "passworttippen" Or $z = "enter" Or $z = "passwort") Then $neu = $z
		; Anmeldung geht weiter: Hinweis zum Selbst-Anmelden weg
		If $selbst And $neu <> "passwort" And $neu <> "passworttippen" Then
			$selbst = False
			ToolTip("")
		EndIf
		If $neu <> "" Then $gemeldet = True
		; Fortschritt: nach jedem neuen Zustand noch mindestens 60 s, damit langsame Seiten (Glooko: Suche,
		; Auswahl, PDF) nicht kurz vor dem letzten Schritt an Zeit= scheitern
		If $neu <> "" And $neu <> $z And $max < TimerDiff($t) + 60000 Then $max = TimerDiff($t) + 60000
		$z = $neu
		Switch $z
			Case ""
				ExitLoop
			Case "passwort"
				; Firefox-Liste der gespeicherten Anmeldungen im Kennungsfeld aufklappen und die erste
				; waehlen; das Skript schickt die Anmeldung ab, sobald das Passwort drinsteht. Die Liste
				; erscheint nicht immer gleich: alle 3 s, hoechstens dreimal; dann den Menschen bitten
				If $liste < 3 Then
					If WinActive($hFF) Then
						$liste += 1
						Send("{DOWN}")
						Sleep(300)
						Send("{ENTER}")
					EndIf
				ElseIf Not $selbst Then
					$selbst = True
					$max = TimerDiff($t) + 180000
					HinweisStehend($s & ": bitte im Firefox anmelden (gespeicherte Anmeldung aus der Liste unter ""Username"" waehlen); " & _
							"danach geht es von selbst weiter")
				EndIf
				$pw = True
			Case "passworttippen"
				; wie "passwort", aber Firefox oeffnet seine Liste erst bei echtem Tippen: das letzte Zeichen
				; der Kennung nachtippen, dann den ersten (darauf gefilterten) Eintrag waehlen
				Local $kennung = Einst($s, "Kennung", "")
				If $tippen < 3 And WinActive($hFF) And $kennung <> "" Then
					$tippen += 1
					; das Skript setzt den Fokus ins Feld
					Sleep(300)
					Send("{END}")
					Send(StringRight($kennung, 1), 1)
					Sleep(700)
					Send("{DOWN}")
					Sleep(300)
					Send("{ENTER}")
				EndIf
				$pw = True
			Case "code"
				; Zwei-Faktor-Anmeldung: der Mensch tippt den Code aus der Mail ein und bestaetigt
				If Not $code Then
					$code = True
					$max = TimerDiff($t) + 300000
					HinweisStehend($s & ": bitte den Code aus der Mail eingeben und bestaetigen (""Bestaetigen und Anmelden"" bzw. ""Fortfahren"")")
				EndIf
			Case "enter"
				; das Skript hat einem Knopf den Fokus gegeben, der nur auf eine echte Benutzeraktion reagiert
				; (Glooko "PDF erstellen"); alle 6 s, hoechstens fuenfmal. Hat ein Popup (z.B. Kontextmenue)
				; den Fokus, ist Firefox nicht aktiv: dann erst wieder aktivieren
				If $enter < 5 Then
					If Not WinActive($hFF) Then
						WinActivate($hFF)
						WinWaitActive($hFF, "", 1)
					EndIf
					If WinActive($hFF) Then
						$enter += 1
						Send("{ENTER}")
					EndIf
				EndIf
			Case "anmelden", "suche", "land", "zweifaktor", "patient"
				If $code Then ToolTip("")
			; Bericht ueber den Druckdialog von Firefox als PDF (EasyView): Druckvorschau und "Speichern unter"
			Case "drucken"
				Return PortalDrucken($s, $hFF, $nach, $vor, $geb, $zeitraum)
			; "gedruckt" kann "fertig" so schnell folgen, dass dieses nie im Titel zu sehen war
			Case "fertig", "gedruckt"
				AblageMerken($s, $nach, $vor, $geb, $zeitraum)
				Return True
			Case "liste"
				Return Meldung("Kein " & $nach & ", " & $vor & " mit Geburtsdatum " & $geb & ", Suchliste angezeigt")
			Case "nichts"
				Return Meldung($s & ": kein Patient " & $nach & " " & $vor)
			Case "keinedaten"
				Return Meldung($s & ": fuer " & $nach & ", " & $vor & " keine Daten")
			Case "ohnekennung"
				Return Meldung($s & ": keine Kennung ([" & $s & "] Kennung= in der INI), bitte selbst anmelden")
			Case "anmeldefehler"
				Return Meldung($s & ": Anmeldung abgelehnt; bitte selbst anmelden (nicht wiederholt probieren, sonst Sperre)")
			Case Else
				Return Meldung($s & ": " & $z)
		EndSwitch
	WEnd
	If $code Or $selbst Then ToolTip("")
	If $pw Then Return Meldung("Bitte im Firefox anmelden (kein gespeichertes Passwort?)")
	If Not $gemeldet Then Return SetExtended(1, False)
	; PDF evtl. doch erstellt (Fertigmeldung verpasst): ein noch erscheinendes trotzdem ablegen
	If $z = "enter" Or $z = "patient" Then AblageMerken($s, $nach, $vor, $geb, $zeitraum)
	Return Meldung($s & ": Zeitueberschreitung" & (($z = "") ? "" : " (zuletzt: " & $z & ")"))
EndFunc

; "drucken": das Benutzerskript hat den Druck der Seite ausgeloest. In der Druckvorschau von Firefox wird
; "Speichern" ([Portal] Speichern= als x,y im Fensterinneren, negativ = vom rechten bzw. unteren Rand) nur
; angeklickt, wenn Firefox zuletzt "Als PDF speichern" benutzt hat (sonst ginge es auf Papier), sonst klickt der
; Mensch; im Fenster "Speichern unter" traegt MOReiter dann Ablage= und Name= (mit {von}, {bis}) ein
Func PortalDrucken($s, $hFF, $nach, $vor, $geb, $zeitraum)
	Local $verz = StringRegExpReplace(Einst($s, "Ablage", ""), "\\$", ""), $name = AblageNeuerName($s, $nach, $vor, $geb)
	If $verz = "" Or $name = "" Then Return Meldung($s & ": [" & $s & "] Ablage= und Name= fehlen")
	Local $ziel = AblageName($verz, NameMitZeitraum($name, $zeitraum))
	Local $dlg = "[REGEXPTITLE:^(Speichern unter|Save As)$; CLASS:#32770]", $h = 0
	Sleep(Int(Einst($s, "DruckWarten", "4000")))
	Local $selbst = Not FirefoxDruckerPdf()
	If Not $selbst Then
		; "Speichern" geht erst, wenn die Vorschau fertig ist: bis zu sechsmal alle 4 s
		Local $p = SmartPixPunkt($hFF, Einst($s, "Speichern", "-272,-121")), $ok = Not @error, $alt = MouseGetPos()
		Opt("MouseCoordMode", 1)
		For $k = 1 To 6
			If Not $ok Or Not WinActive($hFF) Then ExitLoop
			MouseClick("left", $p[0], $p[1], 1, 0)
			MouseMove($alt[0], $alt[1], 0)
			$h = WinWait($dlg, "", 4)
			If $h Then ExitLoop
		Next
		$selbst = Not $h
	EndIf
	If $selbst Then
		HinweisStehend($s & ": bitte in der Druckvorschau Ziel ""Als PDF speichern"" waehlen und ""Speichern"" klicken; " & _
				"Name und Ordner traegt MOReiter dann ein")
		$h = WinWait($dlg, "", 180)
	EndIf
	If Not $h Then
		ToolTip("")
		Return Meldung($s & ": Fenster ""Speichern unter"" nicht erschienen")
	EndIf
	ToolTip("")
	WinActivate($h)
	WinWaitActive($h, "", 3)
	; der Cursor steht beim Oeffnen im Feld "Dateiname" (mit markiertem Vorschlag): wie von Hand alles markieren
	; und den ganzen Pfad einfuegen (ein per ControlSetText gesetzter Text kommt beim Dialog nicht an)
	If Not WinActive($h) Then Return Meldung($s & ": ""Speichern unter"" nicht aktiv, bitte selbst speichern als " & $ziel)
	Sleep(500)
	Send("^a")
	Einfuegen($ziel)
	; zuruecklesen laesst sich das Feld nicht verlaesslich; ob es geklappt hat, zeigt die Datei in der Ablage
	Sleep(500)
	If Not WinActive($h) Then Return Meldung($s & ": ""Speichern unter"" nicht aktiv, bitte selbst speichern als " & $ziel)
	Send("{ENTER}")
	If Not WinWaitClose($h, "", 10) Then Return Meldung($s & ": ""Speichern unter"" blieb offen, bitte selbst speichern als " & $ziel)
	Local $t = TimerInit()
	While Not FileExists($ziel) And TimerDiff($t) < 30000
		Sleep(500)
	WEnd
	If Not FileExists($ziel) Then Return Meldung($s & ": " & $ziel & " nicht angekommen (evtl. unter anderem Namen gespeichert)")
	; fertig geschrieben abwarten (Groesse 2 s unveraendert), dann wie DokimpKurz einordnen
	Local $gr = -1
	$t = TimerInit()
	While TimerDiff($t) < 30000
		Sleep(2000)
		If FileGetSize($ziel) = $gr And $gr > 0 Then ExitLoop
		$gr = FileGetSize($ziel)
	WEnd
	$ziel = DokEinordnen($s, $ziel, $nach, $vor, $geb, $zeitraum)
	Hinweis($s & ": als " & $ziel & " abgelegt")
	Return True
EndFunc

; [Portal] Einordnen=1 (sonst [Dokumente] Einordnen=): ein abgelegtes PDF gleich wie mit DokimpKurz weiter
; einordnen: Sicherheitskopie nach [Dokumente] Eingelesen=\Jahr, Eintrag in quelle (tmbrie, dokprotlist,
; dokprotinh, dokprotvz; Mitarbeiter und Autor = PC-Name), dann nach [Dokumente] Dok=\Pat_id verschieben. Die
; Patientennummer kommt vom Laufzettel-Link oder aus Medical Office (Name und Geburtsdatum eindeutig). Klappt etwas
; nicht, bleibt die Datei in der Ablage (Hinweis, eine schon angelegte Sicherheitskopie wird wieder geloescht);
; liefert den Pfad, unter dem die Datei jetzt liegt
Func DokEinordnen($s, $datei, $nach, $vor, $geb, $zeitraum)
	If Einst($s, "Einordnen", Einst("Dokumente", "Einordnen", "0")) <> "1" Then Return $datei
	Local $pid = $gLinkPid
	If $pid = "" Then
		$pid = MOPatNr($nach, $vor, $geb)
		If @error Then Return DokNicht($s, $datei, "Patientennummer nicht eindeutig (" & $pid & ")")
	EndIf
	Local $name = StringRegExpReplace($datei, "^.*\\", ""), $von = "", $bis = ""
	; Zeitraum des Berichts: sein Ende ist das Quelldatum und bestimmt das Jahr der Sicherheitskopie
	Local $zr = StringRegExp($zeitraum, "^(\d\d)\.(\d\d)\.(\d{4})-(\d\d)\.(\d\d)\.(\d{4})$", 1)
	If Not @error Then
		$von = $zr[2] & "-" & $zr[1] & "-" & $zr[0]
		$bis = $zr[5] & "-" & $zr[4] & "-" & $zr[3]
	EndIf
	Local $svz = StringRegExpReplace(Einst("Dokumente", "Eingelesen", "p:\eingelesen"), "\\$", "") & "\" & (($bis = "") ? @YEAR : StringLeft($bis, 4))
	Local $dvz = StringRegExpReplace(Einst("Dokumente", "Dok", "p:\dok"), "\\$", "") & "\" & $pid
	If Not FileExists($svz) Then DirCreate($svz)
	If Not FileExists($dvz) Then DirCreate($dvz)
	Local $kopie = AblageName($svz, $name)
	If Not FileCopy($datei, $kopie) Then Return DokNicht($s, $datei, "Sicherheitskopie nach " & $svz & " fehlgeschlagen")
	Local $fehler = DokEintragen($s, $pid, $datei, $kopie, $svz, $nach, $vor, $geb, $von, $bis)
	If $fehler <> "" Then
		FileDelete($kopie)
		Return DokNicht($s, $datei, $fehler)
	EndIf
	Local $ziel = AblageName($dvz, $name)
	If Not FileMove($datei, $ziel) Then Return DokNicht($s, $datei, "Verschieben nach " & $dvz & " fehlgeschlagen (eingetragen und gesichert ist sie)")
	Return $ziel
EndFunc

Func DokNicht($s, $datei, $grund)
	MsgBox(48, "MOReiter", $s & ": " & $datei & @CRLF & "nicht eingeordnet: " & $grund, 60)
	Return $datei
EndFunc

; die Eintraege, die DokimpKurz beim Einordnen in quelle macht, alle oder keiner (Transaktion); "" oder der Fehler
Func DokEintragen($s, $pid, $datei, $kopie, $svz, $nach, $vor, $geb, $von, $bis)
	Local $cn = DBVerbinden("quelle")
	If @error Then Return $cn
	$gComFehler = ""
	$cn.BeginTrans
	If $gComFehler <> "" Then Return DokZu($cn, "Datenbankfehler: " & $gComFehler)
	Local $name = StringRegExpReplace($kopie, "^.*\\", ""), $pc = StringLeft(@ComputerName, 15), $t = FileGetTime($datei, 0, 1)
	Local $laend = StringMid($t, 1, 4) & "-" & StringMid($t, 5, 2) & "-" & StringMid($t, 7, 2) & " " & StringMid($t, 9, 2) & ":" & StringMid($t, 11, 2) & ":" & StringMid($t, 13, 2)
	Local $d = StringSplit($geb, ".", 2), $gr = FileGetSize($datei), $jetzt = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC
	Local $p = Int($pid), $l = "'" & $laend & "'"
	; Verzeichnis der Sicherheitskopie
	Local $f = ""
	Local $vz = DokWert($cn, "SELECT id FROM dokprotvz WHERE Verzeichnis='" & SqlText($svz) & "'")
	If @error Then Return DokZu($cn, $vz)
	If $vz = "" Then
		$f = DokSql($cn, "INSERT INTO dokprotvz(Verzeichnis) VALUES('" & SqlText($svz) & "')")
		If $f <> "" Then Return DokZu($cn, $f)
		$vz = DokWert($cn, "SELECT LAST_INSERT_ID()")
		If @error Then Return DokZu($cn, $vz)
	EndIf
	; Behandlungsfall wie DokimpKurz: der zur Aenderungszeit laufende, sonst der letzte davor, sonst der letzte
	$f = DokSql($cn, "INSERT INTO tmbrie(fid,pat_id,tm_pat_id,zeitpunkt,quelldatum,Pfad,Art,NAME,Autor,AktZeit,DokGroe,DokAenD,StByte,qs,qt) VALUES(" & _
			"COALESCE((SELECT MAX(fid) FROM faelle WHERE pat_id=" & $p & " AND " & $l & " BETWEEN bhfb AND bhfe1)," & _
			"(SELECT MAX(fid) FROM faelle WHERE pat_id=" & $p & " AND " & $l & ">bhfb),(SELECT MAX(fid) FROM faelle WHERE pat_id=" & $p & "))," & _
			$p & ",pidtm(" & $p & "),'" & $jetzt & "'," & (($bis = "") ? "NOW()" : "'" & $bis & "'") & ",'" & SqlText($kopie) & "','pdf','" & _
			SqlText($name) & "','" & SqlText($pc) & "',NOW()," & $gr & "," & $l & ",-11,zqsort(" & $l & "),zquart(" & $l & "))")
	If $f <> "" Then Return DokZu($cn, $f)
	$f = DokSql($cn, "INSERT INTO dokprotlist(kPatN,ntum,nImp,urspnm,Patientenname,Gebdat,Ort,Pat_ID,urspName,datName,lAend,groesse,Typ," & _
			"Mitarbeiter,PC,Benutzer,eingetragen,ZielVZID) VALUES(0,1,0,0,'" & SqlText(StringLeft($nach & ", " & $vor, 50)) & "','" & _
			$d[2] & "-" & $d[1] & "-" & $d[0] & "','',"  & $p & ",'" & SqlText($name) & "','" & SqlText($name) & "'," & $l & "," & $gr & _
			",'pdf','" & SqlText($pc) & "','" & SqlText($pc) & "','" & SqlText(StringLeft(@UserName, 15)) & "','" & $jetzt & "'," & $vz & ")")
	If $f <> "" Then Return DokZu($cn, $f)
	Local $id = DokWert($cn, "SELECT LAST_INSERT_ID()")
	If @error Then Return DokZu($cn, $id)
	$f = DokSql($cn, "INSERT INTO dokprotinh(dplbez,nr,Inhalt" & (($bis = "") ? "" : ",von,bis") & ") VALUES(" & $id & ",0,'" & SqlText($s) & "'" & _
			(($bis = "") ? "" : ",'" & $von & "','" & $bis & "'") & ")")
	Return DokZu($cn, $f)
EndFunc

; fuehrt $sql aus; "" oder der Fehler
Func DokSql($cn, $sql)
	$gComFehler = ""
	$cn.Execute($sql)
	Return ($gComFehler = "") ? "" : "Datenbankfehler: " & $gComFehler
EndFunc

; erster Wert der ersten Zeile, "" ohne Zeile; bei Fehler @error und der Fehler
Func DokWert($cn, $sql)
	$gComFehler = ""
	Local $rs = $cn.Execute($sql), $w = ""
	If $gComFehler <> "" Or Not IsObj($rs) Then Return SetError(1, 0, "Datenbankfehler: " & $gComFehler)
	If Not $rs.EOF Then $w = String($rs.Fields(0).Value)
	$rs.Close
	Return $w
EndFunc

; Transaktion abschliessen (ohne Fehler) bzw. zuruecknehmen, Verbindung schliessen
Func DokZu($cn, $fehler)
	$gComFehler = ""
	If $fehler = "" Then
		$cn.CommitTrans
		If $gComFehler <> "" Then $fehler = "Datenbankfehler: " & $gComFehler
	Else
		$cn.RollbackTrans
	EndIf
	$cn.Close
	Return $fehler
EndFunc

; True, wenn Firefox (prefs.js des zuletzt benutzten Profils) zuletzt auf "Als PDF speichern" gedruckt hat
Func FirefoxDruckerPdf()
	Local $basis = @AppDataDir & "\Mozilla\Firefox\Profiles", $such = FileFindFirstFile($basis & "\*"), $neu = "", $zeit = ""
	If $such = -1 Then Return False
	While 1
		Local $d = FileFindNextFile($such)
		If @error Then ExitLoop
		Local $f = $basis & "\" & $d & "\prefs.js"
		If Not FileExists($f) Then ContinueLoop
		If FileGetTime($f, 0, 1) > $zeit Then
			$zeit = FileGetTime($f, 0, 1)
			$neu = $f
		EndIf
	WEnd
	FileClose($such)
	If $neu = "" Then Return False
	Return StringRegExp(FileRead($neu), 'user_pref\("print_printer",\s*"Mozilla Save to PDF"\);')
EndFunc

; bietet Tampermonkey in einem neuen Firefox-Tab an und wartet, bis es eingeschaltet ist
Func AddOnEinrichten()
	Local $pol = TampermonkeyRichtlinie()
	If $pol <> "" Then Return Meldung("Bitte erst als Administrator " & $pol & " loeschen (oder das neue NVerb laufen " & _
			"lassen), Firefox neu starten und Strg+Alt+B wiederholen")
	If Not FirefoxTab(AddOnQuelle()) Then Return False
	HinweisStehend("Einmalig: im Firefox bei ""Tampermonkey hinzufuegen"" auf ""Hinzufuegen"" klicken")
	Local $t = TimerInit()
	While Not TampermonkeyDa()
		If TimerDiff($t) > 180000 Then Return Meldung("Tampermonkey wurde nicht hinzugefuegt")
		Sleep(500)
	WEnd
	BegruessungSchliessen()
	Return True
EndFunc

; nach dem Hinzufuegen oeffnet Tampermonkey seine Begruessungsseite auf tampermonkey.net, die erst um
; Einwilligung in Werbung bittet: diesen Tab (und den von addons.mozilla.org) gleich wieder schliessen.
; Erkannt am Titel des aktiven Tabs; bis 10 s warten, nach dem letzten Schliessen noch 3 s
Func BegruessungSchliessen()
	Local $muster = "[REGEXPTITLE:Mozilla Firefox$; CLASS:MozillaWindowClass]"
	Local $t = TimerInit(), $zu = 0, $letzt = 0
	While TimerDiff($t) < 10000 And $zu < 3
		If $zu > 0 And TimerDiff($letzt) > 3000 Then ExitLoop
		Local $h = WinGetHandle($muster)
		If Not @error And WinActive($h) And StringInStr(FFTitel($h), "Tampermonkey") Then
			Send("^w")
			$zu += 1
			$letzt = TimerInit()
			Sleep(500)
		EndIf
		Sleep(200)
	WEnd
	Sleep(500)
EndFunc

; oeffnet das Benutzerskript in einem neuen Tab (Tampermonkey zeigt dann seine Installationsseite) und
; wartet, bis diese verschwindet ("Installieren": Tampermonkey schliesst den Tab)
; [Aufrufe] Strg+Umsch+T=Skripte: Tampermonkey (falls noch nicht da) und nacheinander die Benutzerskripte aller
; Portale mit Url= in der INI zum Installieren bzw. Aktualisieren anbieten (je "Installieren"/"Aktualisieren"
; klicken oder den Tab schliessen, dann kommt das naechste)
Func SkripteEinrichten()
	If Not ModifierLos() Then Return Meldung("Strg und Umschalt bitte loslassen")
	Local $portale = "", $quelle[2] = [$Ini, $IniVorgaben], $l
	For $datei In $quelle
		$l = IniReadSectionNames($datei)
		If @error Then ContinueLoop
		For $i = 1 To $l[0]
			If Einst($l[$i], "Url", "") = "" Or StringInStr("|" & $portale & "|", "|" & $l[$i] & "|") Then ContinueLoop
			$portale &= ($portale = "" ? "" : "|") & $l[$i]
		Next
	Next
	If $portale = "" Then Return Meldung("Keine Portale mit Url= in der INI")
	If Not TampermonkeyDa() And Not AddOnEinrichten() Then Return False
	Local $p = StringSplit($portale, "|", 2)
	For $k = 0 To UBound($p) - 1
		Hinweis("Benutzerskript " & ($k + 1) & "/" & UBound($p) & ": MOReiter " & $p[$k])
		Sleep(1000)
		If Not SkriptEinrichten($p[$k]) Then Return False
	Next
	Hinweis("Benutzerskripte eingerichtet: " & StringReplace($portale, "|", ", "))
	Return True
EndFunc

Func SkriptEinrichten($s)
	Local $h = FirefoxTab(PortalSkript($s))
	If Not $h Then Return False
	HinweisStehend("Einmalig: im Firefox bei Tampermonkey auf ""Installieren"" klicken")
	; Installationsseite: der erste Titel, der 2 s stehen bleibt und nicht mehr die rohe Skriptadresse ist
	Local $t = TimerInit(), $seite = "", $seit = TimerInit(), $tt
	While TimerDiff($t) < 20000
		$tt = WinGetTitle($h)
		If $tt <> $seite Then
			$seite = $tt
			$seit = TimerInit()
		ElseIf TimerDiff($seit) > 2000 And Not StringInStr($tt, ".user.js") Then
			ExitLoop
		EndIf
		Sleep(200)
	WEnd
	While WinGetTitle($h) = $seite
		If TimerDiff($t) > 180000 Then Return Meldung("Benutzerskript ""MOReiter " & $s & """ wurde nicht installiert")
		Sleep(300)
	WEnd
	Sleep(1000)
	ToolTip("")
	Return True
EndFunc

; oeffnet $adr in einem neuen Firefox-Tab (bzw. startet Firefox damit), liefert das Fenster
; True, wenn zu Firefox ein Dialogfenster (z.B. "Speichern unter") offen ist: Tasten fuer Firefox kaemen dort an
Func FirefoxDialog($h)
	Local $pid = WinGetProcess($h), $l = WinList("[CLASS:#32770]")
	For $i = 1 To $l[0][0]
		If BitAND(WinGetState($l[$i][1]), 2) And WinGetProcess($l[$i][1]) = $pid Then Return True
	Next
	Return False
EndFunc

Func FirefoxTab($adr)
	Local $muster = "[REGEXPTITLE:Mozilla Firefox$; CLASS:MozillaWindowClass]"
	Local $h = WinGetHandle($muster)
	If @error Then
		ShellExecute("firefox.exe", '"' & $adr & '"')
		$h = WinWait($muster, "", 20)
		If Not $h Then Return Meldung("Firefox nicht gefunden")
		Return $h
	EndIf
	WinActivate($h)
	If Not WinWaitActive($h, "", 2) Then Return Meldung("Firefox laesst sich nicht aktivieren")
	If FirefoxDialog($h) Then Return Meldung("Firefox hat noch ein Fenster offen (z.B. ""Speichern unter""), bitte erst schliessen")
	Send("^t")
	Sleep(300)
	Einfuegen($adr)
	Send("{ENTER}")
	Return $h
EndFunc

; ist das Add-on Tampermonkey in einem der Firefox-Profile des Benutzers eingeschaltet? Massgeblich ist
; extensions.json ("active"), die .xpi kann auch von einem abgeschalteten oder entfernten Add-on stammen.
; Profile: alle Ordner unter Profiles\ und alle in profiles.ini genannten (auch ausserhalb, IsRelative=0)
Func TampermonkeyDa()
	Local $basis = @AppDataDir & "\Mozilla\Firefox\", $verz = ""
	Local $such = FileFindFirstFile($basis & "Profiles\*")
	If $such <> -1 Then
		While 1
			Local $d = FileFindNextFile($such)
			If @error Then ExitLoop
			If @extended Then $verz &= "|" & $basis & "Profiles\" & $d
		WEnd
		FileClose($such)
	EndIf
	Local $abschn = IniReadSectionNames($basis & "profiles.ini")
	If Not @error Then
		For $i = 1 To $abschn[0]
			Local $pfad = IniRead($basis & "profiles.ini", $abschn[$i], "Path", "")
			If $pfad = "" Then ContinueLoop
			$pfad = StringReplace($pfad, "/", "\")
			If IniRead($basis & "profiles.ini", $abschn[$i], "IsRelative", "1") = "1" Then $pfad = $basis & $pfad
			$verz &= "|" & $pfad
		Next
	EndIf
	For $v In StringSplit(StringTrimLeft($verz, 1), "|", 2)
		If $v <> "" And TampermonkeyAktiv($v & "\extensions.json") Then Return True
	Next
	Return False
EndFunc

; steht Tampermonkey in dieser extensions.json mit "active":true? Gesucht wird im ganzen Eintrag, bis zum
; naechsten Add-on ({"id":...), da defaultLocale mit langen Mitwirkendenlisten davor stehen kann
Func TampermonkeyAktiv($datei)
	Local $j = FileRead($datei)
	If @error Then Return False
	Local $p = StringInStr($j, '"id":"firefox@tampermonkey.net"')
	If $p = 0 Then Return False
	Local $ende = StringInStr($j, '{"id":"', 0, 1, $p + 10)
	Local $eintrag = ($ende > 0) ? StringMid($j, $p, $ende - $p) : StringMid($j, $p)
	Return StringInStr($eintrag, '"active":true') > 0
EndFunc

; eine Firefox-Unternehmensrichtlinie zu Tampermonkey (distribution\policies.json, frueher von NVerb angelegt)
; kann ein vom Benutzer installiertes Tampermonkey wieder entfernen; liefert ihren Pfad oder ""
Func TampermonkeyRichtlinie()
	Local $vz[3] = [EnvGet("ProgramW6432"), EnvGet("ProgramFiles(x86)"), @ProgramFilesDir]
	For $v In $vz
		If $v = "" Then ContinueLoop
		Local $pol = $v & "\Mozilla Firefox\distribution\policies.json"
		If StringInStr(FileRead($pol), "tampermonkey") Then Return $pol
	Next
	Return ""
EndFunc

; Adresse zum Installieren von Tampermonkey: [Labor] AddOn= (Pfad zu einer .xpi oder Adresse), sonst die
; Seite bei addons.mozilla.org; eine Datei wird als file:-Adresse geoeffnet, Firefox fragt dann nach
Func AddOnQuelle()
	Local $q = Einst("Labor", "AddOn", "https://addons.mozilla.org/de/firefox/addon/tampermonkey/")
	If StringRegExp($q, "^[A-Za-z]:\\") Then Return "file:///" & StringReplace($q, "\", "/")
	If StringLeft($q, 2) = "\\" Then Return "file:" & StringReplace($q, "\", "/")
	Return $q
EndFunc

; %-Kodierung (UTF-8) fuer den Teil hinter "#", ausser Buchstaben, Ziffern und . - _
Func UrlKodiert($s)
	Local $b = StringToBinary($s, 4), $r = ""
	For $i = 1 To BinaryLen($b)
		Local $c = Int(BinaryMid($b, $i, 1))
		If ($c >= 0x30 And $c <= 0x39) Or ($c >= 0x41 And $c <= 0x5A) Or ($c >= 0x61 And $c <= 0x7A) _
				Or $c = 0x2E Or $c = 0x2D Or $c = 0x5F Then
			$r &= Chr($c)
		Else
			$r &= "%" & Hex($c, 2)
		EndIf
	Next
	Return $r
EndFunc

; Link "moreiter:Zucker[:Pat_id];Nachname;Vorname;TT.MM.JJJJ;Markierung|Markierung..." vom Patientenlaufzettel (Klick
; auf eine Zucker-Markierung bzw. Alt+Z fuer alle): wie Strg+Z, aber mit dem Patienten und den Markierungen aus
; dem Link, also ohne Medical Office; die Teile sind percent-codiert (Windows-1252 wie die Seite). Danach wartet
; diese Instanz, bis die abgerufenen PDFs abgelegt sind, und laesst die letzte Meldung noch lesen
Func LinkAufruf($url)
	$gLink = True
	AufrufeLaden()
	Local $r = LinkAusfuehren($url)
	While UBound($gAblagen) > 0
		Sleep(2000)
		AblagenPruefen()
	WEnd
	Sleep(4000)
	Return $r
EndFunc

Func LinkAusfuehren($url)
	; Firefox haengt manchmal einen Schraegstrich an
	Local $t = StringSplit(StringRegExpReplace(StringRegExpReplace($url, "(?i)^moreiter:/*", ""), "/+$", ""), ";", 2)
	If UBound($t) < 5 Then Return Meldung("MOReiter-Link unvollstaendig: " & $url)
	; nur dieser Befehl: eine fremde Seite soll ueber den Link keine Befehlszeilen starten koennen
	; "Zucker:Pat_id" (neuere Laufzettel), sonst sucht DokEinordnen die Patientennummer in Medical Office
	Local $bef = StringRegExp(UrlDekodiert($t[0]), "^Zucker(?::(\d+))?$", 1)
	If @error Then Return Meldung("MOReiter-Link: unbekannter Befehl " & UrlDekodiert($t[0]))
	If UBound($bef) > 0 Then $gLinkPid = String($bef[0])
	$gLinkNach = StringStripWS(UrlDekodiert($t[1]), 3)
	$gLinkVor = StringStripWS(UrlDekodiert($t[2]), 3)
	$gLinkGeb = UrlDekodiert($t[3])
	If $gLinkNach = "" Or Not StringRegExp($gLinkGeb, "^\d\d\.\d\d\.\d{4}$") Then Return Meldung("MOReiter-Link: Name oder Geburtsdatum fehlt: " & $url)
	For $m In StringSplit($t[4], "|", 2)
		Local $w = StringStripWS(UrlDekodiert($m), 3)
		If $w = "" Then ContinueLoop
		ReDim $gLinkMark[UBound($gLinkMark) + 1]
		$gLinkMark[UBound($gLinkMark) - 1] = $w
	Next
	Return Zucker()
EndFunc

; Linkschema "moreiter:" fuer diesen Benutzer auf diese exe (bzw. dieses Skript) eintragen, ohne Adminrechte
Func LinkSchemaEintragen()
	Local $k = "HKCU\Software\Classes\moreiter"
	Local $befehl = '"' & @AutoItExe & '"' & (@Compiled ? "" : ' "' & @ScriptFullPath & '"') & ' "%1"'
	If RegRead($k & "\shell\open\command", "") = $befehl Then Return
	RegWrite($k, "", "REG_SZ", "URL:MOReiter")
	RegWrite($k, "URL Protocol", "REG_SZ", "")
	RegWrite($k & "\shell\open\command", "", "REG_SZ", $befehl)
EndFunc

; %XX-Folgen zurueck in Text, Bytes als Windows-1252
Func UrlDekodiert($s)
	Local $b = Binary(""), $i = 1
	While $i <= StringLen($s)
		Local $h = StringMid($s, $i + 1, 2)
		If StringMid($s, $i, 1) = "%" And StringRegExp($h, "^[0-9A-Fa-f]{2}$") Then
			$b &= BinaryMid(Binary(Int(Dec($h))), 1, 1)
			$i += 3
		Else
			$b &= StringToBinary(StringMid($s, $i, 1), 1)
			$i += 1
		EndIf
	WEnd
	Return BinaryToString($b, 1)
EndFunc

; Nachname, Vorname und Geburtsdatum aus der Seite "Personalien" des geoeffneten Patienten; sie liegt
; auch ohne F6 (unsichtbar unterhalb) im Fenster, sonst wird sie mit F6 kurz geholt
Func PatientDaten($hWnd, ByRef $nach, ByRef $vor, ByRef $geb)
	If $gLink Then
		$nach = $gLinkNach
		$vor = $gLinkVor
		$geb = $gLinkGeb
		Return True
	EndIf
	If PersonalienLesen($hWnd, $nach, $vor, $geb) Then Return True
	WinActivate($hWnd)
	If Not WinWaitActive($hWnd, "", 2) Then Return False
	TasteEinspeisen(0x75, False) ; F6
	Sleep(500)
	Local $ok = PersonalienLesen($hWnd, $nach, $vor, $geb)
	TasteEinspeisen(0x75, False)
	Return $ok
EndFunc

Func PersonalienLesen($hWnd, ByRef $nach, ByRef $vor, ByRef $geb)
	Local $liste = _WinAPI_EnumChildWindows($hWnd, False)
	If @error Then Return False
	$nach = FeldNeben($hWnd, $liste, "&Nachname")
	$vor = FeldNeben($hWnd, $liste, "&Vorname")
	$geb = FeldNeben($hWnd, $liste, "Geb&urtsdatum")
	If $nach = "" Or Not StringRegExp($geb, "^\d{1,2}\.\d{1,2}\.\d{4}$") Then Return False
	; wie im Portal: TT.MM.JJJJ
	Local $d = StringSplit($geb, ".", 2)
	$geb = StringFormat("%02d.%02d.%04d", Int($d[0]), Int($d[1]), Int($d[2]))
	Return True
EndFunc

; Text des Eingabefelds rechts neben der Beschriftung $text (Static) auf gleicher Hoehe
Func FeldNeben($hWnd, $liste, $text)
	For $i = 1 To $liste[0][0]
		If $liste[$i][1] <> "Static" Or _WinAPI_GetWindowText($liste[$i][0]) <> $text Then ContinueLoop
		Local $s = WinGetPos($liste[$i][0])
		If @error Then ContinueLoop
		Local $hEltern = _WinAPI_GetParent($liste[$i][0]), $best = 0, $bestX = 0
		For $j = 1 To $liste[0][0]
			If $liste[$j][1] <> "Edit" Or _WinAPI_GetParent($liste[$j][0]) <> $hEltern Then ContinueLoop
			Local $e = WinGetPos($liste[$j][0])
			If @error Or $e[0] <= $s[0] Or Abs(($e[1] + $e[3] / 2) - ($s[1] + $s[3] / 2)) > 12 Then ContinueLoop
			If $best = 0 Or $e[0] < $bestX Then
				$best = $liste[$j][0]
				$bestX = $e[0]
			EndIf
		Next
		; WM_GETTEXT, das liefert bei Eingabefeldern fremder Programme den Inhalt
		If $best Then Return StringStripWS(ControlGetText($hWnd, "", $best), 3)
	Next
	Return ""
EndFunc

; oeffnet $url im Firefox: im Tab mit dem Portal, falls der aktive oder einer der [Labor] Tabs letzten
; Tabs des vordersten Fensters es zeigt (von rechts her: Strg+9, dann Strg+Bild hoch), sonst in einem
; neuen Tab; laeuft Firefox nicht, wird er damit gestartet
Func FirefoxOeffnen($url, $titel, $s)
	Local $muster = "[REGEXPTITLE:Mozilla Firefox$; CLASS:MozillaWindowClass]"
	; Portalseite (Titel beginnt mit $titel oder lautet "Seite | $titel", z.B. "Patienten | YourLoops") oder eine
	; mit dem Fortschritt des Benutzerskripts
	Local $LaborTitel = "^((?:.* \| )?\Q" & $titel & "\E|MOR:)"
	; dort wartet das Benutzerskript noch (z.B. auf "PDF drucken/speichern") bzw. laeuft der Download: nicht
	; stoeren, statt dessen einen neuen Tab
	Local $belegt = "^MOR:\d+:(fertig|gedruckt)\b"
	Local $h = WinGetHandle($muster)
	If @error Then
		ShellExecute("firefox.exe", '"' & $url & '"')
		$h = WinWait($muster, "", 20)
		If Not $h Then Return Meldung("Firefox nicht gefunden")
		Return $h
	EndIf
	WinActivate($h)
	If Not WinWaitActive($h, "", 2) Then Return Meldung("Firefox laesst sich nicht aktivieren")
	If FirefoxDialog($h) Then Return Meldung("Firefox hat noch ein Fenster offen (z.B. ""Speichern unter""), bitte erst schliessen")
	Local $gefunden = StringRegExp(WinGetTitle($h), $LaborTitel) And Not StringRegExp(WinGetTitle($h), $belegt)
	If Not $gefunden Then
		Local $n = Int(Einst($s, "Tabs", $defLaborTabs)), $vorher = ""
		For $i = 1 To $n
			Send(($i = 1) ? "^9" : "^{PGUP}")
			Sleep(80)
			Local $tt = WinGetTitle($h)
			$gefunden = StringRegExp($tt, $LaborTitel) And Not StringRegExp($tt, $belegt)
			; gefunden, oder Titel unveraendert: ganz links angekommen (bzw. zwei gleiche Tabs, selten)
			If $gefunden Or $tt = $vorher Then ExitLoop
			$vorher = $tt
		Next
	EndIf
	If Not $gefunden Then
		Send("^t")
		Sleep(300)
	EndIf
	; Adresszeile: Strg+L, Adresse, Enter
	Send("^l")
	Sleep(150)
	Einfuegen($url)
	Send("{ENTER}")
	Return $h
EndFunc

; fuegt $txt auf einmal per Zwischenablage ein und stellt deren Text danach wieder her; enthaelt die
; Zwischenablage etwas anderes als Text (z.B. ein Bild), das sonst verloren ginge, wird statt dessen
; schnell getippt
Func Einfuegen($txt)
	Local $alt = ClipGet(), $fehler = @error
	; @error 1 = leer, 2 = kein Text, 3/4 = nicht lesbar
	If $fehler >= 2 Then
		; anderes als Text (Bild, Dateien ...): alle Formate sichern und danach wiederherstellen
		Local $sich = AblageSichern()
		If Not @error Then
			ClipPut($txt)
			Send("^v")
			Sleep(200)
			AblageZurueck($sich)
			Return
		EndIf
		Local $verz = Opt("SendKeyDelay", 0)
		Send($txt, 1)
		Opt("SendKeyDelay", $verz)
		Return
	EndIf
	ClipPut($txt)
	Send("^v")
	; erst wiederherstellen, wenn Firefox eingefuegt hat
	Sleep(200)
	If $fehler = 0 Then ClipPut($alt)
	If $fehler = 1 Then ClipPut("")
EndFunc

; Inhalt der Zwischenablage in allen Formaten mit Speicherblock (Text, Bilder als DIB, Dateilisten, HTML ...)
; als Array [Format, Daten]; GDI-Formate (Bitmap, Metafile, Palette) erzeugt Windows beim Zurueckschreiben aus
; den uebrigen. @error, wenn die Zwischenablage nicht zu oeffnen war
Func AblageSichern()
	Local $erg[0][2]
	Local $r = DllCall("user32.dll", "bool", "OpenClipboard", "hwnd", 0)
	If @error Or Not $r[0] Then Return SetError(1, 0, $erg)
	Local $fmt = 0, $n = 0
	While 1
		$r = DllCall("user32.dll", "uint", "EnumClipboardFormats", "uint", $fmt)
		If @error Or $r[0] = 0 Then ExitLoop
		$fmt = $r[0]
		; CF_BITMAP, CF_METAFILEPICT, CF_PALETTE, CF_ENHMETAFILE, CF_OWNERDISPLAY, CF_DSP...
		If $fmt = 2 Or $fmt = 3 Or $fmt = 9 Or $fmt = 14 Or ($fmt >= 0x80 And $fmt <= 0x8E) Then ContinueLoop
		Local $h = DllCall("user32.dll", "handle", "GetClipboardData", "uint", $fmt)
		If @error Or Not $h[0] Then ContinueLoop
		Local $gr = DllCall("kernel32.dll", "ulong_ptr", "GlobalSize", "handle", $h[0])
		If @error Or $gr[0] = 0 Then ContinueLoop
		Local $p = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", $h[0])
		If @error Or Not $p[0] Then ContinueLoop
		Local $daten = DllStructGetData(DllStructCreate("byte[" & $gr[0] & "]", $p[0]), 1)
		DllCall("kernel32.dll", "bool", "GlobalUnlock", "handle", $h[0])
		ReDim $erg[$n + 1][2]
		$erg[$n][0] = $fmt
		$erg[$n][1] = $daten
		$n += 1
	WEnd
	DllCall("user32.dll", "bool", "CloseClipboard")
	If $n = 0 Then Return SetError(2, 0, $erg)
	Return $erg
EndFunc

Func AblageZurueck($a)
	Local $r = DllCall("user32.dll", "bool", "OpenClipboard", "hwnd", 0)
	If @error Or Not $r[0] Then Return False
	DllCall("user32.dll", "bool", "EmptyClipboard")
	For $i = 0 To UBound($a) - 1
		Local $n = BinaryLen($a[$i][1])
		; 2 = GMEM_MOVEABLE
		Local $h = DllCall("kernel32.dll", "handle", "GlobalAlloc", "uint", 2, "ulong_ptr", $n)
		If @error Or Not $h[0] Then ContinueLoop
		Local $p = DllCall("kernel32.dll", "ptr", "GlobalLock", "handle", $h[0])
		If @error Or Not $p[0] Then ContinueLoop
		DllStructSetData(DllStructCreate("byte[" & $n & "]", $p[0]), 1, $a[$i][1])
		DllCall("kernel32.dll", "bool", "GlobalUnlock", "handle", $h[0])
		; gehoert danach der Zwischenablage
		DllCall("user32.dll", "handle", "SetClipboardData", "uint", $a[$i][0], "handle", $h[0])
	Next
	DllCall("user32.dll", "bool", "CloseClipboard")
	Return True
EndFunc

; Seitentitel ohne " - Mozilla Firefox"
Func FFTitel($h)
	Return StringRegExpReplace(WinGetTitle($h), "\s+\S\s+Mozilla Firefox$", "")
EndFunc

; wartet bis zu $ms auf den Titel "MOR:Nr:Zustand" mit einem anderen Zustand als $ausser, liefert ihn
; oder "" bei Zeitablauf
Func FFZustand($h, $nr, $ausser, $ms)
	Local $t = TimerInit(), $vor = "MOR:" & $nr & ":"
	While TimerDiff($t) < $ms
		Local $tt = FFTitel($h)
		If StringLeft($tt, StringLen($vor)) = $vor Then
			Local $z = StringTrimLeft($tt, StringLen($vor))
			If $z <> $ausser Then Return $z
		EndIf
		Sleep(50)
	WEnd
	Return ""
EndFunc

; Strg+Alt+F1: bildschirmfuellendes Fenster mit der Hilfe; ist es schon offen, schliesst es sich.
; Der Text steht in MOReiter-Hilfe.txt (UTF-8), die beim Kompilieren in die exe eingebettet wird.
; Das Fenster blockiert die Hauptschleife nicht, geschlossen wird es in HilfePruefen.
Func Hilfe()
	If $gHilfe Then Return HilfeWeg()
	Local $datei = @TempDir & "\MOReiter-Hilfe.txt"
	FileInstall("MOReiter-Hilfe.txt", $datei, 1)
	; 256 = UTF-8
	Local $fh = FileOpen($datei, 256)
	If $fh = -1 Then Return Meldung("Hilfetext nicht gefunden")
	Local $txt = FileRead($fh)
	FileClose($fh)
	$txt = StringRegExpReplace($txt, "\r?\n", @CRLF)
	$txt = StringReplace($txt, "{AUFRUFE}", AufrufeText())
	$txt = "Stand " & Stand() & " (" & @AutoItExe & ")" & @CRLF & $txt
	Local $b = @DesktopWidth, $h = @DesktopHeight
	$gHilfe = GUICreate("MOReiter - Hilfe", $b, $h, 0, 0, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetBkColor(0xFFFFFF, $gHilfe)
	; Spalte in der Mitte, so breit wie der Text (ca. 100 Zeichen Consolas 11), hoechstens der Bildschirm
	Local $sb = ($b > 1000) ? 1000 : $b - 20
	Local $ed = GUICtrlCreateEdit($txt, Int(($b - $sb) / 2), 10, $sb, $h - 20, _
			BitOR($ES_MULTILINE, $ES_READONLY, $ES_AUTOVSCROLL, $WS_VSCROLL), 0)
	GUICtrlSetFont($ed, 11, 400, 0, "Consolas")
	GUICtrlSetBkColor($ed, 0xFFFFFF)
	GUISetState(@SW_SHOW, $gHilfe)
	WinActivate($gHilfe)
	; sonst ist der ganze Text markiert
	GUICtrlSendMsg($ed, $EM_SETSEL, 0, 0)
	Return True
EndFunc

; Esc (schliesst GUI-Fenster von selbst) bzw. Schliessen-Ereignis abfragen
Func HilfePruefen()
	Local $msg = GUIGetMsg(1)
	If $msg[0] = $GUI_EVENT_CLOSE And $msg[1] = $gHilfe Then HilfeWeg()
EndFunc

Func HilfeWeg()
	If $gHilfe Then GUIDelete($gHilfe)
	$gHilfe = 0
	Return True
EndFunc

; Strg+Alt+T: Medical Office nach vorn, F6 und Alt+T, bis T oder Strg/Alt losgelassen wird; dann F6
; und zurueck ins vorige Fenster. Die Tasten werden mit $Kennung eingespeist, nachdem Strg und Alt als
; losgelassen gemeldet sind, damit nicht Strg+Alt+F6 ankommt.
Func Termine()
	Local $hVorher = _WinAPI_GetForegroundWindow()
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	ModifierAus()
	If $hVorher <> $hWnd Then
		WinActivate($hWnd)
		If Not WinWaitActive($hWnd, "", 2) Then Return Meldung("Medical Office laesst sich nicht aktivieren")
	EndIf
	Local $ms = Int(Einst("Termine", "Warten", $defTermineWarten))
	TasteEinspeisen(0x75, False) ; F6
	Sleep($ms)
	TasteEinspeisen(0x54, True) ; Alt+T
	; T gilt als gehalten, solange der Hook kein Loslassen gesehen hat; die automatische Wiederholung
	; frischt $gGehaltenZeit auf, ohne sie (verschluckte Loslassen bei Fernsteuerung) nach 1 s Schluss
	While StrgAlt() And $gGehalten = 0x54 And TimerDiff($gGehaltenZeit) < 1000
		Sleep(20)
	WEnd
	TasteEinspeisen(0x75, False) ; F6
	; kam die Wiederholung erst nach mehr als 1 s, hat der Hook sie als neuen Druck vorgemerkt
	If $gAufgabe = "Termine" Then $gAufgabe = ""
	If $hVorher <> $hWnd And WinExists($hVorher) Then
		Sleep($ms)
		WinActivate($hVorher)
	EndIf
	Return True
EndFunc

; schickt die Taste $vk (mit Alt) als eigene, vom Hook durchgelassene Eingabe
Func TasteEinspeisen($vk, $mitAlt)
	; 2 = KEYEVENTF_KEYUP
	If $mitAlt Then DllCall("user32.dll", "none", "keybd_event", "byte", 0xA4, "byte", 0, "dword", 0, "ulong_ptr", $Kennung)
	DllCall("user32.dll", "none", "keybd_event", "byte", $vk, "byte", 0, "dword", 0, "ulong_ptr", $Kennung)
	DllCall("user32.dll", "none", "keybd_event", "byte", $vk, "byte", 0, "dword", 2, "ulong_ptr", $Kennung)
	If $mitAlt Then DllCall("user32.dll", "none", "keybd_event", "byte", 0xA4, "byte", 0, "dword", 2, "ulong_ptr", $Kennung)
EndFunc

; Strg+Alt+U: die Liste im Container Karteikarte aktivieren, damit Pfeiltasten usw. dort wirken;
; ist keine sichtbar (z.B. auf dem Reiter ePAAbr), erst zum Krankenblatt wechseln
Func Unten()
	Local $hWnd, $hCtrl
	If Not HolLeiste($hWnd, $hCtrl) Then Return False
	WinActivate($hWnd)
	WinWaitActive($hWnd, "", 2)
	Local $hGrid = KarteiListe($hWnd)
	If Not $hGrid Then
		If Not Reiter("Krankenblatt") Then Return False
		Sleep(Int(Einst("Filter", "Warten", $defFilterWarten)))
		$hGrid = KarteiListe($hWnd)
		If Not $hGrid Then Return Meldung("Keine Karteikarte gefunden")
	EndIf
	If Einst("Karteikarte", "Art", "Fokus") <> "Klick" Then
		ControlFocus($hWnd, "", $hGrid)
		If ControlGetHandle($hWnd, "", ControlGetFocus($hWnd)) = $hGrid Then Return True
	EndIf
	; Fokus hat nicht gegriffen oder Klick gewuenscht: in die markierte Zeile klicken, sonst in die
	; oberste, jeweils ganz links in die Datumsspalte (dort liegen keine Symbole)
	Local $g = WinGetPos($hGrid)
	If @error Then Return Meldung("Karteikarte nicht sichtbar")
	Opt("PixelCoordMode", 1)
	Local $y = $g[1] + $KbKopfHoehe + 8
	Local $m = PixelSearch($g[0] + 4, $g[1] + $KbKopfHoehe, $g[0] + 4, $g[1] + $g[3] - 3, $KbMarkiertFarbe, 6)
	If Not @error Then $y = $m[1] + 4
	Return KlickPunkt($hWnd, $g[0] + 20, $y, "Karteikarte verdeckt")
EndFunc

; die Liste im Krankenblatt-Container: die groesste sichtbare, die mit dem Fokus nur, wenn sie
; mindestens halb so gross ist (die Filterliste links ist ebenfalls ein TmoStringGrid in der
; Karteikarte und darf nicht gewinnen, nur weil sie den Fokus hat); MO haelt mehrere TKarteikarteForm
; vor, die sichtbare ist nicht immer die erste, z.B. nach einem Wechsel in ein anderes Programm und zurueck
Func KarteiListe($hWnd)
	Local $hFokus = ControlGetHandle($hWnd, "", ControlGetFocus($hWnd))
	If @error Or Not KarteiGrid($hFokus) Then $hFokus = 0
	; nur sichtbare Fenster, auch die Eltern muessen sichtbar sein
	Local $liste = _WinAPI_EnumChildWindows($hWnd, True)
	If @error Then Return 0
	Local $best = 0, $bestFlaeche = 0
	For $i = 1 To $liste[0][0]
		If $liste[$i][1] <> "TmoStringGrid" Or Not KarteiGrid($liste[$i][0]) Then ContinueLoop
		Local $p = WinGetPos($liste[$i][0])
		If @error Then ContinueLoop
		If $p[2] * $p[3] > $bestFlaeche Then
			$best = $liste[$i][0]
			$bestFlaeche = $p[2] * $p[3]
		EndIf
	Next
	If $hFokus And $hFokus <> $best Then
		Local $f = WinGetPos($hFokus)
		If Not @error And $f[2] * $f[3] * 2 >= $bestFlaeche Then Return $hFokus
	EndIf
	Return $best
EndFunc

; sichtbare Liste einer Karteikarte?
Func KarteiGrid($h)
	Return _WinAPI_GetClassName($h) = "TmoStringGrid" And _WinAPI_IsWindowVisible($h) _
			And _WinAPI_GetClassName(_WinAPI_GetParent($h)) = "TfrmKarteikarteControl"
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
; Geklickt wird mit der echten Maus, Alt vorher als losgelassen gemeldet (ModifierAus): auf einen per Nachricht
; geschickten Klick (ControlClick) beginnt die Bereichsauswahl ein Ziehen (Halteverbots-Mauszeiger),
; statt den Eintrag zu waehlen, und die Liste rechts bekaeme den Tastaturfokus nicht.
Func BereichWechsel($richtung)
	Local $hWnd = WinGetHandle("[ACTIVE]")
	If _WinAPI_GetClassName($hWnd) <> $TagesKlasse Then Return False
	Local $hGrid = ControlGetHandle($hWnd, "", "[CLASS:TNewStringGrid; INSTANCE:1]")
	If @error Then Return Meldung("Liste der Tagesuebersicht nicht gefunden")
	; Alt als losgelassen melden, damit die Klicks ohne Alt ankommen; so kann Alt fuer mehrere
	; Schritte gehalten bleiben
	ModifierAus()
	Local $eintr, $akt, $mx, $h
	Local $fehler = BereichSuchen($hWnd, $hGrid, $eintr, $akt, $mx, $h)
	If $fehler <> "" Then Return Meldung($fehler)
	Local $g = WinGetPos($hGrid)
	If @error Then Return False
	Local $n = UBound($eintr), $ziel = Mod($akt + $richtung + $n, $n)
	; oberste Zeile rechts merken, um zu sehen, wann die Liste neu aufgebaut ist
	Local $zy = $g[1] + $KopfHoehe
	Local $vorher = PixelChecksum($g[0] + 2, $zy, $g[0] + $g[2] - 20, $zy + $ZeilenHoehe - 1)
	Opt("MouseCoordMode", 1)
	Local $alt = MouseGetPos()
	MouseClick("left", $mx, $eintr[$ziel] + Int($h / 2), 1, 0)
	Local $t0 = TimerInit(), $max = Int(Einst("Tagesuebersicht", "BereichWarten", $defBereichWarten))
	While TimerDiff($t0) < $max
		Sleep(50)
		If PixelChecksum($g[0] + 2, $zy, $g[0] + $g[2] - 20, $zy + $ZeilenHoehe - 1) <> $vorher Then ExitLoop
	WEnd
	; bis die Zeile fertig gezeichnet ist
	Sleep(100)
	Local $nameX = $g[0] + Int(Einst("Tagesuebersicht", "NameX", $defNameX)), $zm = $zy + Int($ZeilenHoehe / 2)
	; leere Liste: dort ist nur weisse Flaeche
	If PixelGetColor($nameX, $zm) <> 0xFFFFFF Then MouseClick("left", $nameX, $zm, 1, 0)
	MouseMove($alt[0], $alt[1], 0)
	Return True
EndFunc

; sucht in der Bereichsauswahl links neben der Liste $hGrid den gelb umrandeten Eintrag und zaehlt die
; uebrigen: $eintr = Oberkanten aller Eintraege, $akt = Stelle des gewaehlten, $mx = Mitte x, $h = Hoehe
; eines Eintrags; Rueckgabe "" bei Erfolg, sonst der Grund
Func BereichSuchen($hWnd, $hGrid, ByRef $eintr, ByRef $akt, ByRef $mx, ByRef $h)
	Local $g = WinGetPos($hGrid)
	If @error Then Return "Liste der Tagesuebersicht nicht sichtbar"
	Local $pt = DllStructCreate("int X;int Y")
	_WinAPI_ClientToScreen($hWnd, $pt)
	; Bereich links neben der Liste
	Local $x1 = $pt.X, $x2 = $g[0] - 1, $y1 = $g[1], $y2 = $g[1] + $g[3] - 1
	If $x2 - $x1 < 40 Then Return "Bereichsauswahl links nicht sichtbar"
	Local $farbe = Int(Einst("Tagesuebersicht", "BereichRahmen", $defBereichRahmen))
	Opt("PixelCoordMode", 1)
	If Not BildLesen($x1, $y1, $x2, $y2) Then Return "Bildschirm nicht lesbar"
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
		Return "Kein gelb umrandeter Eintrag links gefunden"
	EndIf
	; Rahmenhoehe; der Eintrag ist 2 Pixel groesser nach jeder Seite, dazu 1 Pixel Abstand
	Local $fu = $fy
	While $fu < $fy + 60 And FarbeNah($fx, $fu + 1, $farbe, 10)
		$fu += 1
	WEnd
	$h = $fu - $fy + 5
	$mx = Int(($fx + $rx) / 2)
	Local $raster = $h + 1, $oben = $fy - 2, $bx = $fx - 1
	; Oberkanten aller Eintraege, der gewaehlte an Stelle $akt
	Local $e[1] = [$oben]
	$eintr = $e
	$akt = 0
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
	Return ""
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
	MouseClick("left", $g[0] + Int(Einst("Tagesuebersicht", "BereichX", $defBereichX)), $mitte, 1, 0)
	; Kontextmenue (Offen / In Arbeit / erledigt) abwarten
	If Not MenueWarten(True) Then
		MouseMove($alt[0], $alt[1], 0)
		Return Meldung("Menue Offen/In Arbeit/erledigt ist nicht erschienen")
	EndIf
	Send("e")
	MenueWarten(False)
	Sleep(Int(Einst("Tagesuebersicht", "Warten", $defAbhakenWarten)))
	; nachrueckende Zeile markieren; war es die letzte Zeile, ist dort jetzt leere (weisse) Flaeche,
	; dann die daruber
	Local $nameX = $g[0] + Int(Einst("Tagesuebersicht", "NameX", $defNameX))
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
	Local $hKnopf = LetztePatKnopf($hWnd)
	If Not $hKnopf Then Return Meldung("Pfeil neben der Patientensuche nicht gefunden")
	Return KnopfKlick($hWnd, $hKnopf)
EndFunc

Func LetztePatKnopf($hWnd)
	Local $hSuche = ControlGetHandle($hWnd, "", "[CLASS:TAlnumEdit; INSTANCE:1]")
	If @error Then Return 0
	Return RandKnopf(_WinAPI_GetParent($hSuche), "rechts")
EndFunc

; drei Striche links oben: der oberste, linkeste Knopf im Kopfbereich
Func Menue()
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $hKnopf = MenueKnopf($hWnd)
	If Not $hKnopf Then Return Meldung("Menueknopf nicht gefunden")
	Return KnopfKlick($hWnd, $hKnopf)
EndFunc

Func MenueKnopf($hWnd)
	Local $hKopf = ControlGetHandle($hWnd, "", "[CLASS:THeaderForm; INSTANCE:1]")
	If @error Then Return 0
	Return RandKnopf($hKopf, "linksoben")
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
	If Einst("Allgemein", "EchteMaus", "0") = "1" Then
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
	Local $reiter = Einst("Filter", "Reiter", $defFilterReiter)
	If Not ReiterAktiv($reiter, $hWnd, $hCtrl) Then
		If Not Reiter($reiter) Then Return False
		Sleep(Int(Einst("Filter", "Warten", $defFilterWarten)))
	EndIf
	Local $x = Int(Einst("Filter", "X", $defFilterX))
	Local $y = Int(Einst("Filter", "Y", $defFilterY)) + $i * Number(Einst("Filter", "Abstand", $defFilterAbstand))
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
	If Einst("Allgemein", "EchteMaus", "0") = "1" Then
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
		Local $abst = Round(($y - Int(Einst("Filter", "Y", $defFilterY))) / $i, 2)
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
	$hCtrl = ControlGetHandle($hWnd, "", Einst("Allgemein", "Control", $defCtrl))
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

	If Einst("Allgemein", "EchteMaus", "0") = "1" Then
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
			$x = Int(Einst("Kartei", "X", $defKarteiX))
			$y = Int(Einst("Kartei", "Y", $defKarteiY))
		Case "Krankenblatt"
			$x = Int(Einst("Krankenblatt", "X", $defKrbX))
			$y = Int(Einst("Krankenblatt", "Y", $defKrbY))
		Case "ePA"
			$x = Int(Einst("ePA", "X", $defEpaX))
			$y = Int(Einst("ePA", "Y", $defEpaY))
		Case "ePAAbr"
			$x = Int(Einst("ePAAbr", "X", $defEpaAbrX))
			$y = Int(Einst("ePAAbr", "Y", $defEpaAbrY))
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
	Local $farbe = Int(Einst("Allgemein", "AktivFarbe", $defAktivFarbe))
	Local $tol = Int(Einst("Allgemein", "Toleranz", $defToleranz))
	Opt("PixelCoordMode", 1)
	PixelSearch($p[0] + $x - 12, $p[1] + $y - 5, $p[0] + $x + 12, $p[1] + $y + 5, $farbe, $tol)
	Return Not @error
EndFunc

; merkt sich die aktuelle Mausposition relativ zur Reiterleiste
Func Kalibrieren($name)
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return MOStarten()
	Local $hCtrl = ControlGetHandle($hWnd, "", Einst("Allgemein", "Control", $defCtrl))
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
	$gMeldung = $txt
	SmartPixLog("Meldung: " & $txt)
	If $CmdLine[0] > 0 And Not $gLink Then
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

; wie Hinweis, bleibt aber stehen, bis ToolTip("") oder der naechste Hinweis ihn ersetzt
Func HinweisStehend($txt)
	AdlibUnRegister("HinweisWeg")
	ToolTip($txt, Default, Default, "MOReiter")
EndFunc

Func NeustartTaste()
	$gNeustart = True
EndFunc

; Strg+Alt+Esc: eine neue Instanz starten (wartet, bis diese beendet ist) und diese beenden. Ausserdem den
; Hook auch waehrend einer laufenden Aufgabe alle 5 s neu einhaengen (sonst nur in der Hauptschleife)
Func NeustartPruefen()
	If TimerDiff($gHookAdlib) > 5000 Then
		HookErneuern()
		$gHookAdlib = TimerInit()
	EndIf
	If Not $gNeustart Then Return
	$gNeustart = False
	ToolTip("MOReiter startet neu", Default, Default, "MOReiter")
	Local $befehl = '"' & @AutoItExe & '"' & (@Compiled ? "" : ' "' & @ScriptFullPath & '"') & " Neustart"
	If Not Run($befehl, @ScriptDir) Then Return Meldung("Neustart fehlgeschlagen")
	Exit
EndFunc

Func HinweisWeg()
	AdlibUnRegister("HinweisWeg")
	ToolTip("")
EndFunc
