#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=MOReiter.ico
#AutoIt3Wrapper_Outfile=MOReiter.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Kompilieren mit Icon: Aut2exe.exe /in MOReiter.au3 /out MOReiter.exe /icon MOReiter.ico
; MOReiter: waehlt in Medical Office den Reiter "Kartei" oder "Krankenblatt" per Mausklick
; Aufruf:  MOReiter.exe Kartei | Krankenblatt   -> einmal klicken und beenden (z.B. aus VB6 per Shell)
;          MOReiter.exe Wechsel                  -> Kartei, bzw. Krankenblatt, wenn Kartei schon aktiv ist
;          MOReiter.exe                          -> bleibt resident mit Hotkeys:
;            Strg+Alt+K        Kartei, ist Kartei schon aktiv: Krankenblatt
;            Strg+Alt+L        Krankenblatt
;            Strg+Alt+Umsch+K  Kalibrieren: Maus auf Reiter "Kartei" halten und druecken
;            Strg+Alt+Umsch+L  Kalibrieren: Maus auf Reiter "Krankenblatt" halten und druecken
;            Beenden ueber das Tray-Menue (kein Strg+Alt+Q, das waere AltGr+Q = @)
; Die Reiterleiste ist ein Delphi-Control (TmoTabSet) ohne eigene Handles je Reiter,
; deshalb wird relativ zur linken oberen Ecke dieses Controls geklickt.
#include <Misc.au3>
Opt("MustDeclareVars", 1)
Opt("WinTitleMatchMode", 4)

Const $MOFenster = "[REGEXPTITLE:^Medical Office; CLASS:OWL_Window]"
; INI im Benutzerprofil, weil das Programmverzeichnis unter "Program Files" nicht beschreibbar ist
Const $Ini = @AppDataDir & "\MOReiter\MOReiter.ini"

; Voreinstellungen (Pixel relativ zum Control), werden durch Kalibrieren in der INI ueberschrieben
Const $defCtrl = "[CLASS:TmoTabSet; INSTANCE:1]"
Const $defKarteiX = 20, $defKarteiY = 10, $defKrbX = 85, $defKrbY = 10
; Hintergrund des aktiven Reiters (hellblau, gemessen RGB 210,226,247), inaktive Reiter sind weiss
Const $defAktivFarbe = 0xD2E2F7, $defToleranz = 12

If $CmdLine[0] > 0 Then
	Exit Reiter($CmdLine[1]) ? 0 : 1
EndIf
If _Singleton("MOReiter_resident", 1) = 0 Then Exit ; laeuft schon

; Icon aus der Datei daneben laden, falls es beim Kompilieren nicht in die exe gekommen ist
If FileExists(@ScriptDir & "\MOReiter.ico") Then TraySetIcon(@ScriptDir & "\MOReiter.ico")
TraySetToolTip("MOReiter: Strg+Alt+K Kartei/Wechsel, Strg+Alt+L Krankenblatt")
HotKeySet("^!k", "Wechsel")
HotKeySet("^!l", "Krankenblatt")
HotKeySet("^!+k", "KalibKartei")
HotKeySet("^!+l", "KalibKrankenblatt")
While 1
	Sleep(100)
WEnd

Func Wechsel()
	Reiter("Wechsel")
EndFunc

Func Krankenblatt()
	Reiter("Krankenblatt")
EndFunc

Func KalibKartei()
	Kalibrieren("Kartei")
EndFunc

Func KalibKrankenblatt()
	Kalibrieren("Krankenblatt")
EndFunc

; liefert True bei Erfolg
Func Reiter($name)
	Local $x, $y
	If $name <> "Wechsel" And Not HolPos($name, $x, $y) Then Return Meldung("Unbekannter Reiter: " & $name)
	Local $hWnd = WinGetHandle($MOFenster)
	If @error Then Return Meldung("Medical Office nicht gefunden")
	Local $ctrl = IniRead($Ini, "Allgemein", "Control", $defCtrl)
	Local $hCtrl = ControlGetHandle($hWnd, "", $ctrl)
	If @error Or Not BitAND(WinGetState($hCtrl), 2) Then Return Meldung("Reiterleiste nicht sichtbar (Patient geoeffnet?)")
	If $name = "Wechsel" Then
		$name = KarteiAktiv($hWnd, $hCtrl) ? "Krankenblatt" : "Kartei"
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
		Case Else
			Return False
	EndSwitch
	Return True
EndFunc

; prueft am Bildschirm, ob der Reiter "Kartei" hellblau hinterlegt, also aktiv ist
; (eine PixelSearch ueber ein kleines Rechteck dauert nur wenige Millisekunden)
Func KarteiAktiv($hWnd, $hCtrl)
	Local $x, $y
	HolPos("Kartei", $x, $y)
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
	If @error Then Return Meldung("Medical Office nicht gefunden")
	Local $hCtrl = ControlGetHandle($hWnd, "", IniRead($Ini, "Allgemein", "Control", $defCtrl))
	If @error Then Return Meldung("Reiterleiste nicht gefunden")
	Opt("MouseCoordMode", 1)
	Local $m = MouseGetPos(), $p = WinGetPos($hCtrl)
	Local $x = $m[0] - $p[0], $y = $m[1] - $p[1]
	If $x < 0 Or $y < 0 Or $x >= $p[2] Or $y >= $p[3] Then Return Meldung("Maus steht nicht auf der Reiterleiste")
	DirCreate(@AppDataDir & "\MOReiter")
	IniWrite($Ini, $name, "X", $x)
	IniWrite($Ini, $name, "Y", $y)
	TrayTip("MOReiter", $name & " kalibriert: " & $x & ", " & $y, 3)
	Return True
EndFunc

Func Meldung($txt)
	If $CmdLine[0] > 0 Then
		ConsoleWriteError($txt & @CRLF)
	Else
		TrayTip("MOReiter", $txt, 3, 2)
	EndIf
	Return False
EndFunc
