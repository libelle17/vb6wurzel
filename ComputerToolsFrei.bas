Attribute VB_Name = "ComputerToolsFrei"
' notwendig für dieses Modul: Verweis auf Microsoft VBScript Regular Expressions: c:\windows\syswow64\vbscript.dll\3
Option Explicit
Public Const ap1$ = "-u administrator -p so"
Declare Function GetComputerName& Lib "kernel32" Alias "GetComputerNameA" (ByVal lbbuffer$, nSize&)
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer$, nSize&)
Public Const p1$ = "so"
Public Const doalsAd$ = "pstools\psexec.exe "
Public Const acceu$ = "/accepteula "
Public Const ap2$ = "nne13"
Public Declare Function GetFileAttributes& Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName$)


Public Function ODBCStr$()
 Static ergstr$
' SELECT CASE CptName
'  Case "SZN1", "SZN2", "ANMELDR", "ANMELDRNEU", "ANMELDR1", "SZO1", "SZS1", "SR", "SONO1", "MITTE1", "BZW1", "ANMELDL", "LABOR1", "SR2", "SR3", "SONO2"
   
   If ergstr = "" Then ergstr = maxODBC
   ODBCStr = ergstr
'   ODBCStr = "MySQL ODBC 5.3 Unicode Driver"

'  Case "MITTE", "ANMELD1", "SPN1", "BZW2", "HSS"
'   ODBCStr = "MySQL ODBC 5.2 Unicode Driver"
'  Case Else
'   ODBCStr = "MySQL ODBC 5.1 Driver"
'   ODBCStr = "MySQL ODBC 5.2 Unicode Driver"
'  END SELECT
End Function  ' ODBCStr

Public Function maxODBC()
  Dim i%, arrValueNames, arrValueTypes
  Dim objRegistry As Object, strKeyPath$
  Dim Version#, maxv#
  Const strComputer$ = "."
  Const HKEY_LOCAL_MACHINE = &H80000002
  Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
  strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
  objRegistry.EnumVALUES HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
  For i = 0 To UBound(arrValueNames)
   Dim rEx As RegExp
   Set rEx = New RegExp
   rEx.Pattern = "MySQL ODBC [0123456789]\.[0123456789]* Unicode Driver"
   If rEx.test(arrValueNames(i)) Then
    Version = Mid$(arrValueNames(i), 12, 3)
    If Version > maxv Then
     maxv = Version
     maxODBC = arrValueNames(i)
    End If ' Version > maxv Then
   End If ' rEx.Test(arrValueNames(i)) Then
  Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in maxODBC/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' maxODBC()

Function CptName$() ' Computername in Großbuchstaben
 Static Cpt$
 If Cpt <> "" Then
  CptName = Cpt
 Else
 Dim sTxt As String * 64
  On Error GoTo fehler
  Call GetComputerName(sTxt, 64)
  CptName = UCase$(sTxt)
  If InStr(sTxt, vbNullChar) > 1 Then
   CptName = Left$(sTxt, InStr(sTxt, vbNullChar) - 1)
  End If
  Cpt = UCase$(CptName) ' ucase 12.8.15, zur Sicherheit
 End If
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in CptName/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' CptName

Public Function UserName$()
    Dim Cn$
    Dim ls&
    Dim res&
    On Error GoTo fehler
    Cn = String$(1024, 0)
    ls = 1024
    res = GetUserName(Cn, ls)
    If res <> 0 Then
        UserName = Mid$(Cn, 1, InStr(Cn, vbNullChar) - 1)
    Else
        UserName = ""
    End If
  Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UserName/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Username


Public Function getVerzAnteil(Datei$, Optional GrenzBchst$ = "\", Optional rest$)
 Dim i%, aktl%, lenge%
 lenge = Len(Datei)
 For i = lenge To 0 Step -1
  If Mid$(Datei, i, 1) = GrenzBchst Then
   aktl = i
   Exit For
  End If
 Next i
 getVerzAnteil = Left$(Datei, i)
 rest = Right$(Datei, lenge - i + 1)
End Function ' getVerzAnteil

Function DirKorr$(Datei$, Optional Dwc$)
     Dim D2$, i%
     Dwc = Datei
     If InStrB(Datei, "Â") <> 0 Then
      DirKorr = REPLACE$(Datei, "+Â", "ö")
      Dwc = REPLACE$(Dwc, "+Â", "?Â")
     End If
     If InStrB(Datei, "++") <> 0 Then
      DirKorr = REPLACE$(Datei, "++", "ü")
      Dwc = REPLACE$(Dwc, "++", "??")
     End If
     If InStrB(Datei, "+ñ") <> 0 Then
      DirKorr = REPLACE$(Datei, "+ñ", "ä")
      Dwc = REPLACE$(Dwc, "+ñ", "?ñ")
     End If
     If InStrB(Datei, "+ƒ") <> 0 Then
      DirKorr = REPLACE$(Datei, "+ƒ", "ß")
      Dwc = REPLACE$(Dwc, "+ƒ", "?ƒ")
     End If
     If InStrB(Datei, "+û") <> 0 Then
      DirKorr = REPLACE$(Datei, "+û", "Ö")
      Dwc = REPLACE$(Dwc, "+û", "?û")
     End If
     If InStrB(Datei, "+ä") <> 0 Then
      DirKorr = REPLACE$(Datei, "+ä", "Ä")
      Dwc = REPLACE$(Dwc, "+ä", "?ä")
     End If
     If Dwc = Datei Then
       For i = 1 To Len(Datei)
        DirKorr = DirKorr & Chr$(Asc(Mid$(Datei, i, 1)))
       Next i
       If DirKorr = Datei Then
'        Stop
       Else
'        Stop
       End If
     End If
End Function ' DirKorr
'Function SplitNeu&(ByVal q$, Sep$, erg$(), Optional nichtWenn$, Optional Bis$) ' da Split() Speicher fraß
'' in einem Fragment darf nicht nichtWenn enthalten sein, ohne dass Bis enthalten ist
' Dim p1&, p2&, Slen&, lSlen&, obExit%, runde&, p3&, p4&, obgesprungen%
' On Error GoTo fehler
' IF NOT ISNULL(q) THEN
'  Slen = Len(Sep)
'  For runde = 1 To 2
'   p2 = 1
'   lSlen = 0
'   Do
'    obgesprungen = 0
'    p1 = p2
'    p2 = InStr(p1 + lSlen, q, Sep)
'    IF p2 <> 0 AND LenB(nichtWenn) <> 0 THEN
'     p3 = InStr(p1 + lSlen, q, nichtWenn)
'     IF p3 <> 0 AND p3 < p2 THEN
'      p4 = InStr(p3 + 1, q, Bis)
'      IF p4 <> 0 AND p4 > p2 THEN
'       p2 = InStr(p4, q, Sep)
'       obgesprungen = True
'      END IF
'     END IF
'    END IF
'    IF p2 = 0 THEN p2 = Len(q) + 1: obExit = True
'    IF runde = 2 THEN
'     erg(SplitNeu) = Mid$(q, p1 + lSlen, p2 - p1 - lSlen)
'     IF obgesprungen THEN
'      erg(SplitNeu) = replace$(erg(SplitNeu), Sep, " ")
'     END IF
'    END IF
'    SplitNeu = SplitNeu + 1
'    lSlen = Slen
'    IF obExit THEN Exit Do
'   Loop
'   IF runde = 1 THEN
'    ReDim erg(SplitNeu - 1)
'    SplitNeu = 0
'    obExit = 0
'   END IF
'  Next runde
' END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in SplitNeu/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' SplitNeu ' aufSplit
'Function stest()
' Dim erg&, ergeb$(), Sep
' Sep = Array("hie", "da")
' erg = SplitNeu("hie und da gibt hie oder da Sinn", "hier", ergeb)
'' erg = SplitNeuArr("hie und da gibt hie oder da Sinn", sep, ergeb)
'' ergeb = Split("hie und da", "dort")
' Stop
'End Function

Function SplitNeuArr&(ByRef q$, Sep, erg$(), Optional obmitsep%) ' da Split() Speicher fraß; sep = Array aus Trennstrings; ob SEPARATOR auch noch im Splitstring stehen soll
 Dim p1&, p2&, Slen&(), lSlen&, obExit%, runde&, aktSep%, p2min&, gew%
 On Error GoTo fehler
 If Not IsNull(q) Then
  ReDim Slen(LBound(Sep) To UBound(Sep))
  For aktSep = LBound(Sep) To UBound(Sep)
   Slen(aktSep) = Len(Sep(aktSep))
  Next aktSep
  For runde = 1 To 2
   p2min = 1
   lSlen = 0
   Do
    p1 = p2min
    p2min = 0
    For aktSep = LBound(Sep) To UBound(Sep)
     p2 = InStr(p1 + lSlen, q, Sep(aktSep))
     If p2 <> 0 And ((p2 < p2min) Or (p2min = 0)) Then gew = aktSep: p2min = p2
    Next aktSep
    If p2min = 0 Then p2min = Len(q) + 1: obExit = True
    If runde = 2 Then
     If obmitsep Then
      erg(SplitNeuArr) = Mid$(q, p1, p2min - p1)
     Else
      erg(SplitNeuArr) = Mid$(q, p1 + lSlen, p2min - p1 - lSlen)
     End If
    End If
    SplitNeuArr = SplitNeuArr + 1
    lSlen = Slen(gew)
    If obExit Then Exit Do
   Loop
   If runde = 1 Then
    ReDim erg(SplitNeuArr - 1)
    SplitNeuArr = 0
    obExit = 0
   End If
  Next runde
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitNeuArr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SplitNeuArr, aus aufSplit
'Function zztest()
' Dim sql$, rAF&
' sql = "INSERT INTO zz(v1,v2,i2) VALUES('bc','b',33)"
' Lese.ProgStart
'' myEFrag sql, rAF
' InsKorr DBCn, sql, rAF
' Debug.Print rAF
'End Function
' Führt SQL-Befehl aus
' Handelt es sich um einen singulären insert-Befehl und ist ein Tabellenfeld zu klein, so wird es vergrößert
' nur sinnvoll, falls vergrößerbare Felder enthalten sind

Public Function ZwischenStr$(ByRef ST$, ByRef s1$, ByRef s2$)
 Dim p1&, p2&
 p1 = InStr(ST, s1)
 If p1 <> 0 Then
  p1 = p1 + Len(s1)
  p2 = InStr(p1, ST, s2)
  ZwischenStr = Mid$(ST, p1, p2 - p1)
 End If
End Function ' ZwischenStr

'Function CurDB$(DBCn)
' Dim cn As New ADODB.Connection
' IF VarType(DBCn) = vbString AND NOT IsObject(DBCn) THEN
'  cn.Open DBCn
''  CurDB = cn.Properties("Current Catalog"), 8.4.10 ist identisch mit:
'  CurDB = cn.DefaultDatabase
'  IF LenB(CurDB) = 0 OR CurDB = "null" THEN CurDB = DefDB(cn)
' Else
''  CurDB = DBCn.Properties("Current Catalog"), 8.4.10 ist identisch mit:
'  CurDB = DBCn.DefaultDatabase
'  IF LenB(CurDB) = 0 OR CurDB = "null" THEN CurDB = DefDB(DBCn)
' END IF
'End Function
'
'Function DefDB$(DBCn)
' Dim spos&, sp2&, runde%, dWort$
' IF Not DBCn Is Nothing THEN
'  IF DBCn.State = 0 THEN
'   For runde = 1 To 2
'    IF runde = 1 THEN dWort = "data source=" ELSE dWort = "database="
'    spos = InStr(LCase$(DBCn), dWort)
'    IF spos <> 0 THEN
'     sp2 = InStr(spos, DBCn, ";")
'     IF sp2 = 0 THEN sp2 = Len(DBCn)
'     DefDB = Mid$(DBCn, spos + Len(dWort), sp2 - spos - Len(dWort))
'     IF (Left$(DefDB, 1) = """" AND Right$(DefDB, 1) = """") OR (Left$(DefDB, 1) = "'" AND Right$(DefDB, 1) = "'") THEN DefDB = Mid$(DefDB, 2, Len(DefDB) - 2)
'     Exit For
'    END IF
'   Next runde
'  Else
'   DefDB = DBCn.DefaultDatabase
'   On Error Resume Next
'   IF LenB(DefDB) = 0 THEN DefDB = DBCn.Properties("Data Source Name").Value
'  END IF
' END IF
' IF LenB(DefDB) = 0 THEN DefDB = "quelle"
'End FUNCTION ' DefDB
'
' wird in DateiLese benötigt
Function GetSvr$(DBCn)
 Dim spos&, sp2&
 spos = InStr(1, DBCn, "server=", vbTextCompare)
 If spos <> 0 Then
  sp2 = InStr(spos, DBCn, ";")
  If sp2 = 0 Then sp2 = Len(DBCn)
  GetSvr = Mid$(DBCn, spos + 7, sp2 - spos - 7)
 End If
End Function ' GetSvr
'
'Function fUmwfSQL(q$, Optional obMy% = True) AS CString ' flexibles Umwandeln für SQL
' Const Maxz% = 2
' Dim pos&, obumw%, zwi$, z$(Maxz), vz$(Maxz), j%
' z(0) = "'"
' z(1) = "\"
' z(2) = """"
' IF obMy THEN
'  vz(0) = "\"
'  vz(1) = "\"
'  vz(2) = "\"
' Else
'  vz(0) = "'"
'  vz(1) = "\"
'  vz(2) = vNS
' END IF
' SET fUmwfSQL = New CString
' fUmwfSQL = q
' For pos = fUmwfSQL.Length To 1 Step -1
'  For j = 0 To Maxz
'   IF fUmwfSQL.Mid(pos, 1) = z(j) THEN
'    obumw = 0
'    IF pos = 1 THEN
'     obumw = True
'    ElseIf fUmwfSQL.Mid(pos - 1, 1) <> vz(j) THEN
'     obumw = True
'    END IF
'    IF obumw THEN
'     zwi = fUmwfSQL.Mid(pos)
'     fUmwfSQL.Cut (pos - 1)
'     fUmwfSQL.AppVar Array(vz(j), zwi)
'    Else
'     pos = pos - 1
'    END IF
'    j = Maxz
'   END IF
'  Next j
' Next pos
' IF InStrB(fUmwfSQL, Chr$(0)) <> 0 THEN ' aus umw in tabübertr
'  fUmwfSQL = replace$(fUmwfSQL, Chr$(0), vNS)
' END IF
'End FUNCTION ' fumwfsql

' entfernt endständige Kommandozeilenparameter und Anführungszeichen
' kommt vor in getIViewPfad
Function GetExeF$(q$)
 Dim sp$(), i%, j%, TS$
 On Error GoTo fehler
' IF FSO Is Nothing THEN SET FSO = CreateObject("scripting.filesystemobject")
 sp = Split(REPLACE(q, """", ""))
 For i = UBound(sp) To 0 Step -1
  TS = ""
  For j = 0 To i
   TS = TS + IIf(j = 0, "", " ") + sp(j)
  Next j
'  IF fileexists(TS) THEN
'  IF LenB(dir(TS)) <> 0 THEN
  If FileExists(TS) Then
   GetExeF = TS
   Exit Function
  End If
 Next i
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetExeF/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetExeF

#If ohnewsh = 0 Then
Public Function getIViewPfad$()
 On Error GoTo fehler:
 If getIViewPfad = "" Then getIViewPfad = GetExeF(getReg(2, "Software\Classes\Applications\i_view64.exe\shell\open\command", ""))
 If getIViewPfad = "" Then getIViewPfad = GetExeF(getReg(2, "Software\Classes\Applications\i_view32.exe\shell\open\command", ""))
Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in getIViewPfad/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getIViewPfad
#End If

' da Funktion dir( nicht "einttrittsinvariant")
Public Function DirExists%(ByVal DirName$)
 Dim dwAtt&
 dwAtt = GetFileAttributes(DirName)
 If dwAtt = -1 Then ' INVALID_FILE_ATTRIBUTES
 ElseIf CBool(dwAtt And &H10) Then ' FILE_ATTRIBUTE_DIRECTORY
  DirExists = True
 End If
End Function ' DirExists

' da Funktion dir( nicht "einttrittsinvariant")
Public Function FileExists%(ByVal FName$)
 Dim dwAtt&
 dwAtt = GetFileAttributes(FName)
 If dwAtt = -1 Then ' INVALID_FILE_ATTRIBUTES
 ElseIf CBool(dwAtt And &H10) Then ' FILE_ATTRIBUTE_DIRECTORY
 Else
  FileExists = True
 End If
End Function ' DirExists

Public Function zeigan&(Datei$, Optional modus% = vbMaximizedFocus)
 Static pdat$
 On Error GoTo fehler
 Const uvd$ = "\notepad++\notepad++.exe"
 If pdat = "" Then
  pdat = "c:\Program Files" & uvd
  If Dir(pdat) = "" Then
   pdat = "c:\Program Files (x86)" & uvd
   If Dir(pdat) = "" Then pdat = ""
  End If
 End If
 If pdat <> "" Then
  On Error Resume Next
  zeigan = Shell("""" & pdat & """ """ & Datei & """", modus)
  If Err.Number <> 0 Then Err.Clear: Shell """" & Datei & """", modus
  If Err.Number <> 0 Then MsgBox ("Fehler beim Ausführen von: '" & pdat & " """ & Datei & """" & "'")
  On Error GoTo fehler
 End If
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), "", CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in zeigan/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' zeigan

' Beliebige Datei auslesen und
' Inhalt als String zurückgeben
Public Function ReadFile$(ByVal sFilename As String)
  Dim F%
  Dim sInhalt$
  ' Prüfen, ob Datei existiert
  If Dir$(sFilename, vbNormal) <> "" Then
    ' Datei im Binärmodus öffnen
    F = FreeFile: Open sFilename For Binary As #F
    ' Größe ermitteln und Variable entsprechend
    ' mit Leerzeichen füllen
    sInhalt = Space$(LOF(F))
    ' Gesamten Inhalt in einem "Rutsch" einlesen
    Get #F, , sInhalt
    ' Datei schliessen
    Close #F
  End If
  ReadFile = sInhalt
End Function ' ReadFile

Public Function syscmd(art%, Optional Inhalt$)
 On Error Resume Next
 Select Case art
  Case 4 ' acSysCmdSetStatus
   Forms(0).Fuß = Inhalt
  Case 5 ' acSysCmdClearStatus
   Forms(0).Fuß = vNS
 End Select
' Debug.Print Inhalt
 Err.Clear
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in syscmd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' syscmd(art%, Optional Inhalt$)

