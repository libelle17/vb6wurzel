Attribute VB_Name = "FensterZu"
' 6.11.26: schliesst den von FCAufruf geoeffneten FreeCommander nach einer Stunde wieder.
' FCAufruf selbst endet wie bisher sofort; das Warten uebernimmt ein unsichtbarer zweiter Lauf "/zu <Prozess> <Sekunden>".
' Der laeuft aus einer Kopie in %TEMP%, damit die Exe in den Programmen nicht gesperrt ist und NVerb sie ersetzen kann.
' Schliesst der Benutzer den FreeCommander vorher, endet der Waechter mit ihm.
' Nicht geschlossen wird, solange der Benutzer gerade darin arbeitet (Fenster vorn und Eingabe in den letzten 5 min) oder
' solange ein anderes Fenster des Prozesses offen ist (z.B. Kopierfortschritt, Betrachter, Einstellungen); dann erneut
' nach einer Minute. Meldungen nur mit einem Knopf (z.B. "Zeitueberschreitung beim Pruefen, ob ... existiert: P:\", die
' sich bei kurz nicht erreichbarem P: stapeln und das Hauptfenster sperren) werden dabei weggeklickt.
' Analog in ExpAufruf\FensterZu.bas (dort Explorer-Fenster).
Option Explicit
Public Const SchliessenNachSek& = 3600
Private Const PruefAbstandMs& = 60000, MeldungAbstandMs& = 2000, ArbeitetMs& = 300000
Private Const WM_CLOSE& = &H10, SYNCHRONIZE& = &H100000, WAIT_TIMEOUT& = &H102
Private Type LASTINPUTINFO
 cbSize As Long
 dwTime As Long
End Type
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpVon$, ByVal lpNach$, ByVal bNichtUeberschreiben&) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwZugriff&, ByVal bErben&, ByVal dwPid&) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hObjekt&, ByVal dwMs&) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObjekt&) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (ByRef lii As LASTINPUTINFO) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hEltern&, ByVal hNach&, ByVal lpKlasse$, ByVal lpTitel$) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd&, ByRef pid&) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd&) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd&) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpKlasse$, ByVal nMax&) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&) As Long

' startet den Waechter fuer den Prozess pid, Fehler duerfen den Aufruf nicht stoeren
Sub StarteWaechter(ByVal pid&)
 Dim tv$, d$, Kopie$, Exe$
 On Error Resume Next
 If pid = 0 Then Exit Sub
 tv = Environ$("TEMP") & "\"
 ' alte Kopien aufraeumen (die noch laufenden lassen sich nicht loeschen und bleiben)
 d = Dir$(tv & App.EXEName & "_*.exe")
 Do While Len(d) > 0
  Kill tv & d
  d = ""
  d = Dir$
 Loop
 Exe = App.Path & "\" & App.EXEName & ".exe"
 Kopie = tv & App.EXEName & "_" & pid & ".exe"
 If CopyFile(Exe, Kopie, 0) = 0 Then Kopie = Exe
 Call Shell(Chr$(34) & Kopie & Chr$(34) & " /zu " & pid & " " & SchliessenNachSek, vbHide)
 Call Prot("Waechter fuer Prozess " & pid & ": " & Kopie & IIf(Err.Number = 0, "", " Fehler: " & Err.Description))
End Sub ' StarteWaechter

' zweiter Lauf: par = "<Prozess> <Sekunden>"
Sub WaechterLauf(par$)
 Dim t$(), pid&, sek&, hProz&, warte&
 On Error GoTo fehler
 t = Split(Trim$(par), " ")
 pid = Val(t(0))
 sek = Val(t(1))
 hProz = OpenProcess(SYNCHRONIZE, 0, pid)
 If hProz = 0 Then Exit Sub
 warte = sek * 1000
 Do While WaitForSingleObject(hProz, warte) = WAIT_TIMEOUT
  warte = SchliessVersuch(pid)
 Loop
 Call CloseHandle(hProz)
 Call Prot("Waechter: Prozess " & pid & " beendet")
 Exit Sub
fehler:
 Call Prot("Waechter, Fehler " & Err.Number & ": " & Err.Description)
End Sub ' WaechterLauf

' ein Versuch, gibt die Wartezeit in ms bis zum naechsten zurueck
Function SchliessVersuch&(ByVal pid&)
 Dim h&, wem&, kl$, haupt&, fremd As Boolean, nMeld&, meld&(99), i&
 SchliessVersuch = PruefAbstandMs
 h = FindWindowEx(0, 0, vbNullString, vbNullString)
 Do While h <> 0
  Call GetWindowThreadProcessId(h, wem)
  If wem = pid Then
   If IsWindowVisible(h) <> 0 Then
    kl = Klasse(h)
    If Left$(kl, 13) = "FreeCommander" Then
     haupt = h
    ElseIf kl = "TApplication" Then ' unsichtbares Delphi-Anwendungsfenster
    ElseIf kl = "#32770" And NurEinKnopf(h) And nMeld <= UBound(meld) Then
     meld(nMeld) = h
     nMeld = nMeld + 1
    Else
     fremd = True
    End If
   End If
  End If
  h = FindWindowEx(0, h, vbNullString, vbNullString)
 Loop
 Call GetWindowThreadProcessId(GetForegroundWindow(), wem)
 If wem = pid And BenutzerAktiv() Then
  Call Prot("Waechter: Benutzer arbeitet im Prozess " & pid)
  Exit Function
 End If
 If fremd Then
  Call Prot("Waechter: anderes Fenster im Prozess " & pid & " offen")
  Exit Function
 End If
 If nMeld > 0 Then
  Call Prot("Waechter: " & nMeld & " Meldung(en) im Prozess " & pid & " wegklicken")
  For i = 0 To nMeld - 1
   Call PostMessage(meld(i), WM_CLOSE, 0, 0)
  Next i
  SchliessVersuch = MeldungAbstandMs
  Exit Function
 End If
 If haupt <> 0 Then
  If IsWindowEnabled(haupt) <> 0 Then
   Call Prot("Waechter: schliesse Prozess " & pid)
   Call PostMessage(haupt, WM_CLOSE, 0, 0)
  End If
 End If
End Function ' SchliessVersuch

' Meldung mit genau einem Knopf (schliesst wie OK); Fragen und Fortschrittsdialoge haben mehr oder andere Elemente
Private Function NurEinKnopf(ByVal h&) As Boolean
 Dim k&
 k = FindWindowEx(h, 0, "Button", vbNullString)
 If k = 0 Then Exit Function
 If FindWindowEx(h, k, "Button", vbNullString) <> 0 Then Exit Function
 If FindWindowEx(h, 0, "Edit", vbNullString) <> 0 Then Exit Function
 NurEinKnopf = True
End Function ' NurEinKnopf

Private Function Klasse$(ByVal h&)
 Dim puf$, l&
 puf = String$(128, 0)
 l = GetClassName(h, puf, Len(puf))
 Klasse = Left$(puf, l)
End Function ' Klasse

' Eingabe (Maus, Tastatur) in den letzten 5 Minuten; nachts bleibt ein vorderes Fenster also nicht offen
Private Function BenutzerAktiv() As Boolean
 Dim lii As LASTINPUTINFO
 On Error Resume Next ' GetTickCount laeuft nach knapp 25 Tagen ins Negative
 BenutzerAktiv = True
 lii.cbSize = Len(lii)
 If GetLastInputInfo(lii) <> 0 Then BenutzerAktiv = (GetTickCount() - lii.dwTime < ArbeitetMs)
End Function ' BenutzerAktiv
