Attribute VB_Name = "InsKorrMod"
' Klassenmodul CString.cls nötig
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds&)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public maxlz& ' maximale Laufzahl für DBCnOpen
Public DBCn As New ADODB.Connection
Public DBCnS$ ' Connection-String von DBCn, da auf Vista dieser unvollständig => immer mitführen
Public ErrNumber&, ErrDescr$, ErrSource$, ErrLastDllError&
Public obTrans% ' ob BeginTrans für DBCn aufgerufen wurde => in
' Public DefaultDatabase$

Public Enum ShellSpecialFolderConstants
  ssfDESKTOP = &H0                   ' <Desktop>
  ssfPROGRAMS = &H2                  ' Startmenü\Programme
  ssfPERSONAL = &H5                  ' Eigene Dateien
  ssfFAVORITES = &H6                 ' <Benutzer>\Favoriten
  ssfSTARTUP = &H7                   ' Startmenü\Programme\Autostart
  ssfRECENT = &H8                    ' <Benutzer>\Recent
  ssfSENDTO = &H9                    ' <Benutzer>\SendTo
  ssfSTARTMENU = &HB                 ' <Benutzer>\Startmenü
  ssfDESKTOPDIRECTORY = &H10         ' <Benutzer>\Desktop
  ssfNETHOOD = &H13                  ' <Benutzer>\Netzwerkumgebung
  ssfFONTS = &H14                    ' Windows\Fonts
  ssfTEMPLATES = &H15                ' <Benutzer>\Vorlagen
  ssfCOMMONSTARTMENU = &H16          ' All Users\Startmenü
  ssfCOMMONPROGRAMS = &H17           ' All Users\Startmenü\Programme
  ssfCOMMONSTARTUP = &H18            ' All Users\Startmenü\Autostart
  ssfCOMMONDESKTOPDIRECTORY = &H19   ' All Users\Desktop
  ssfAPPDATA = &H1A                  ' <Benutzer>\Anwendungsdaten
  ssfPRINTHOOD = &H1B                ' <Benutzer>\Druckumgebung
  ssfCOOKIES = &H21                  ' <Benutzer>\Cookies
  ssfHISTORY = &H22                  ' <Benutzer>\Lokale Einstell.\Verlauf
  ssfCOMMONTEMPLATES = &H2D          ' All Users\Vorlagen
  ssfCOMMONDOCUMENTS = &H2E          ' All Users\Dokumente
End Enum
Private Type SHITEMID
  cb   As Long
  abID As Byte
End Type


Private Type ITEMIDLIST
  mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation& Lib "Shell32" (ByVal hwndOwner&, ByVal nFolder&, ByRef ppidl As ITEMIDLIST)
Private Declare Function SHGetPathFromIDList& Lib "Shell32" (ByVal pidList&, ByVal lpBuffer$)
Private Const S_OK = 0
Private Const MAX_PATH = 260

Public Function GetSpecialFolder$(ByVal Folder As ShellSpecialFolderConstants)
  Dim tIIDL As ITEMIDLIST
  Dim strPath$
  If SHGetSpecialFolderLocation(0, Folder, tIIDL) = S_OK Then
    strPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(tIIDL.mkid.cb, strPath) <> 0 Then
      GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    End If
  End If
End Function ' GetSpecialFolder(ByVal Folder As ShellSpecialFolderConstants) As String

Function CurDB$(DBCn)
 Dim Cn As New ADODB.Connection
 On Error GoTo fehler
 If VarType(DBCn) = vbString And Not IsObject(DBCn) Then
  Cn.Open DBCn
'  CurDB = cn.Properties("Current Catalog"), 8.4.10 ist identisch mit:
  CurDB = Cn.DefaultDatabase
  If LenB(CurDB) = 0 Or CurDB = "null" Then CurDB = DefDB(Cn)
 Else
'  CurDB = DBCn.Properties("Current Catalog"), 8.4.10 ist identisch mit:
  CurDB = DBCn.DefaultDatabase
  If LenB(CurDB) = 0 Or CurDB = "null" Then CurDB = DefDB(DBCn)
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in CurDB/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' CurDB$(DBCn)

'Function testtrans()
' Dim rs As New ADODB.Recordset, rAF&
' Lese.ProgStart
' myEFrag ("use haerzte_neu")
' If Not Lese.obMySQL Then
'' myEFrag ("begin")
' else
'' DBCn.BeginTrans
' end if
' myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF
' SET rs = myEFrag("SHOW VARIABLES WHERE variable_name = 'autocommit'", rAF)
' Debug.Print rs!Value
' SET rs = Nothing
' myFrag rs, "SELECT COUNT(0) zl FROM ort"
' Debug.Print rs!zl
' SET rs = Nothing
' myEFrag ("INSERT INTO ort(ort)VALUES('Neuhofen11')")
' myFrag rs, "SELECT COUNT(0) zl FROM ort"
' Debug.Print rs!zl
'
' myEFrag ("commit")
'' DBCn.CommitTrans
' SET rs = Nothing
' myFrag rs, "SELECT COUNT(0) zl FROM ort"
' Debug.Print rs!zl
' SET rs = myEFrag("SHOW VARIABLES WHERE variable_name = 'autocommit'", rAF)
' Debug.Print rs!Value
'End Function

Function DefDB$(DBCn)
 Dim spos&, sp2&, runde%, dWort$
 On Error GoTo fehler
 If Not DBCn Is Nothing Then
  If DBCn.State = 0 Then
   For runde = 1 To 2
    If runde = 1 Then dWort = "data source=" Else dWort = "database="
    spos = InStr(1, DBCn, dWort, vbTextCompare)
    If spos <> 0 Then
     sp2 = InStr(spos, DBCn, ";")
     If sp2 = 0 Then sp2 = Len(DBCn)
     DefDB = Mid$(DBCn, spos + Len(dWort), sp2 - spos - Len(dWort))
     If (Left$(DefDB, 1) = """" And Right$(DefDB, 1) = """") Or (Left$(DefDB, 1) = "'" And Right$(DefDB, 1) = "'") Then DefDB = Mid$(DefDB, 2, Len(DefDB) - 2)
     Exit For
    End If
   Next runde
  Else
   DefDB = DBCn.DefaultDatabase
   On Error Resume Next
   If LenB(DefDB) = 0 Then
    DefDB = DBCn.Properties("Data Source Name").Value
    If InStrB(DefDB, "") <> 0 Then DefDB = vNS
   End If
  End If
 End If
 If LenB(DefDB) = 0 Then
  Select Case LCase$(App.EXEName)
   Case "dateilese"
    DefDB = "quelle"
  End Select
 End If ' LenB(DefDB) = 0 THEN
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DefDB/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' DefDB

Function GetSvr$(DBCn)
 Dim spos&, sp2&
 On Error GoTo fehler
 spos = InStr(1, DBCn, "server=", vbTextCompare)
 If spos <> 0 Then
  sp2 = InStr(spos, DBCn, ";")
  If sp2 = 0 Then sp2 = Len(DBCn)
  GetSvr = Mid$(DBCn, spos + 7, sp2 - spos - 7)
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetSvr/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetServer

Function SplitNeu&(ByRef q$, Sep$, erg$(), Optional nichtWenn$, Optional Bis$) ' da Split() Speicher fraß
' split
' in einem Fragment darf nicht nichtWenn enthalten sein, ohne dass Bis enthalten ist
 Dim p1&, p2&, Slen&, lSlen&, obExit%, runde&, p3&, p4&, obgesprungen%
 On Error GoTo fehler
 If Not IsNull(q) Then
  Slen = Len(Sep)
  For runde = 1 To 2
   p2 = 1
   lSlen = 0
   Do
    obgesprungen = 0
    p1 = p2
    p2 = InStr(p1 + lSlen, q, Sep)
    If p2 <> 0 And LenB(nichtWenn) <> 0 Then
     p3 = InStr(p1 + lSlen, q, nichtWenn)
     If p3 <> 0 And p3 < p2 Then
      p4 = InStr(p3 + 1, q, Bis)
      If p4 <> 0 And p4 > p2 Then
       p2 = InStr(p4, q, Sep)
       obgesprungen = True
      End If
     End If
    End If
    If p2 = 0 Then p2 = Len(q) + 1: obExit = True
    If runde = 2 Then
     erg(SplitNeu) = Mid$(q, p1 + lSlen, p2 - p1 - lSlen)
     If obgesprungen Then
      If Sep <> " " Then
       erg(SplitNeu) = REPLACE$(erg(SplitNeu), Sep, " ")
      End If
     End If
    End If
    SplitNeu = SplitNeu + 1
    lSlen = Slen
    If obExit Then Exit Do
   Loop
   If runde = 1 Then
    ReDim erg(SplitNeu - 1)
    SplitNeu = 0
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SplitNeu/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SplitNeu ' aufSplit

Public Function fUmwfSQL(q$, Optional obmy% = True) As CString ' flexibles Umwandeln für SQL
 Const Maxz% = 2
 Dim pos&, obumw%, zwi$, z$(Maxz), vz$(Maxz), j%
 On Error GoTo fehler
 z(0) = "'"
 z(1) = "\"
 z(2) = """"
 If obmy Then
  vz(0) = "\"
  vz(1) = "\"
  vz(2) = "\"
 Else
  vz(0) = "'"
  vz(1) = "\"
  vz(2) = vNS
 End If
 Set fUmwfSQL = New CString
 fUmwfSQL = q
 For pos = fUmwfSQL.length To 1 Step -1
  For j = 0 To Maxz
   If fUmwfSQL.Mid(pos, 1) = z(j) Then
    obumw = 0
    If pos = 1 Then
     obumw = True
    ElseIf fUmwfSQL.Mid(pos - 1, 1) <> vz(j) Then
     obumw = True
    End If
    If obumw Then
     zwi = fUmwfSQL.Mid(pos)
     fUmwfSQL.Cut (pos - 1)
     fUmwfSQL.AppVar Array(vz(j), zwi)
    Else
     pos = pos - 1
    End If
    j = Maxz
   End If
  Next j
 Next pos
 If InStrB(fUmwfSQL, Chr$(0)) <> 0 Then ' aus umw in tabübertr
  fUmwfSQL = REPLACE$(fUmwfSQL, Chr$(0), vNS)
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in fUmwfSQL/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' fumwfsql

Sub InsKorr(Cn As ADODB.Connection, sql$, Optional rAF&, Optional ErrDes$, Optional restarttrans%, Optional ErrNr&)
 Dim Feld$, UFELD$, Tbl$, p1$, p2$, spl1s$, spl2s$, s1$(), s2$(), csql As New CString, ix&, i&, j&
 Dim cDB$, svr$, CNs$
 Dim rs As New ADODB.Recordset
' Dim rErr As New ADODB.Recordset
 Dim altmode$, obneuMode%, raM As New ADODB.Recordset
 Dim Dtl$(5)
 Dim altDes$
 Dim altErrDes$, FMeld$
 Dim keinetrans%, keinfehler%
 Const insAnw$ = "INSERT INTO"
 CNs = Cn.Properties![extended properties]
 On Error GoTo fehler
 FNr = 1
 If InStrB(CNs, "MySQL") <> 0 Or InStrB(CNs, "MSDASQL") <> 0 Then
  myFrag raM, "SELECT @@session.sql_mode", adOpenStatic, Cn, adLockReadOnly
  If Not raM.BOF Then
   altmode = raM.Fields(0)
   If InStrB(altmode, "strict_trans_tables") = 0 Then
    obneuMode = True
    myEFrag "SET SESSION sql_mode='strict_trans_tables'", rAF, Cn
   End If ' InStrB(altMode, "strict_trans_tables") = 0 Then
  End If ' not raM.BOF
 End If ' InStrB(CNs, "MySQL") <> 0 Or InStrB(CNs, "MSDASQL") <> 0 Then
 FNr = 2
anfang:
 On Error Resume Next
 For j = 1 To 2
  FNr = 2 + j
  rAF = 0
nochmal:
  myEFrag sql, rAF, Cn, True, ErrNr, ErrDes
'  ErrNr = Err.Number
'  ErrDescr = Err.Description
  
  If ErrNr <> 0 Then
'   Set rErr = Nothing
'   myFrag rErr, "SHOW ERRORS", adOpenStatic, Cn, adLockReadOnly
'   If rErr.BOF Then
'    ErrDes = vNS
'   Else
'    ErrDes = rErr!Message
'   End If
   If j = 1 And ErrNr = -2147467259 And InStrB(ErrDescr, "Daten zu lang") = 0 And InStrB(ErrDescr, "Data too long") = 0 Then ' -2147467259 ' [MySQL][ODBC 5.1 Driver][mysqld-5.1.32-log]Cannot add OR UPDATE a child row: a FOREIGN KEY constraint fails
    If altDes = ErrDescr And ErrDes = altErrDes Then
     FMeld = "Fehler:" & vbCrLf & altDes & vbCrLf & altErrDes & vbCrLf & "bei:" & vbCrLf & "nicht behebbar!"
'     MsgBox FMeld
     Ausgeb FMeld
     syscmd 4, FMeld
     Err.Raise 17
     GoTo exyt
    End If
    altDes = ErrDescr
    altErrDes = ErrDes
    If InStrB(ErrDescr, "'READ-COMMITTED'") <> 0 Then
     myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAF, Cn
    End If
   Else
'   IF rAF <> 0 THEN Stop
    Exit For
   End If
  Else
   Exit For
  End If
 Next j
 FNr = 5
 If ErrNr <> 0 Then
  If LenB(ErrDes) = 0 Then
   Debug.Print "Fehler in Inskorr: " & ErrDescr
  Else
   Debug.Print "Fehler in Inskorr: " & ErrDescr & vbCrLf & "       " & ErrDes
  End If
  On Error GoTo fehler
  Dtl(0) = "Data too long for column"
  Dtl(1) = "Daten zu lang f?r Feld"
  Dtl(2) = "Daten zu lang für Feld" ' hypothetisch
  Dtl(3) = "Out of range value for column"
  Dtl(4) = "Falscher decimal-Wert"
'  Exit Sub
  Dtl(5) = "Das Feld ist zu klein für die Datenmenge" ' Access, noch nicht ausprogrammiert
'  IF InStrB(ErrDes, "Falscher integer-Wert") <> 0 THEN Stop
  cDB = CurDB(Cn)
  svr = GetSvr(Cn)
  FNr = 6
  For j = 0 To 5 ' 5 = z.B. Access
   FNr = 10 + j
   Dim obErrDes%
   obErrDes = (InStrB(ErrDes, Dtl(j)) <> 0)
   If obErrDes Then
    If InStrB(CNs, "MySQL") <> 0 Or InStrB(CNs, "MSDASQL") <> 0 Then
     FNr = 105
     If j = 4 Then Dtl(j) = "r Feld"
     FNr = 106
     p1 = InStr(ErrDes, Dtl(j)) + Len(Dtl(j)) + 2
     FNr = 107
     p2 = InStr(p1, ErrDes, "'") - p1
     FNr = 108
     If p2 < 0 Then GoTo exyt
     Feld = Mid$(ErrDes, p1, p2)
     FNr = 109
     UFELD = UCase$(Feld)
     FNr = 110
    Else ' Access: alle Felder
     FNr = 111
     Feld = vNS
     FNr = 112
     UFELD = vNS
     FNr = 113
    End If
    FNr = 20 + j
    If InStrB(sql, insAnw) <> 0 Then
     p1 = InStr(sql, insAnw) + Len(insAnw) + 1
     p2 = InStr(p1, sql, "(") - p1
     Tbl = Trim$(Mid$(sql, p1, p2))
     If Left$(Tbl, 1) = "`" And Right$(Tbl, 1) = "`" Then Tbl = Mid$(Tbl, 2, Len(Tbl) - 2)
     p1 = InStr(sql, "(") + 1
     p2 = InStr(p1, sql, ")") - p1 ' wenn keine Klammern im Namen vorkommen
     spl1s = Mid$(sql, p1, p2)
     p2 = CLng(p2) + CLng(p1)
     p1 = InStr(p2, sql, "(") + 1
     p2 = Len(sql)
'     Do
'      IF Mid$(sql, p2, 1) = ")" THEN Exit Do
'     Loop
     FNr = 30 + j
     p2 = p2 - p1
 '    p2 = InStr(p1, sql, ")") - p1
     spl2s = Mid$(sql, p1, p2)
     SplitNeu spl1s, ",", s1, "`", "`"
     SplitNeu spl2s, ",", s2, "'", "'"
     If LenB(UFELD) <> 0 Then
      For i = 0 To UBound(s1)
       s1(i) = Trim$(s1(i))
       If Left$(s1(i), 1) = "`" Then
        If Right$(s1(i), 1) = "`" Then
         If UCase$(Mid$(s1(i), 2)) = UFELD & "`" Then
          ix = i
          Exit For
         End If
        End If
       ElseIf UCase$(s1(i)) = UFELD Then
        ix = i
        Exit For
       End If
      Next i
     End If
     FNr = 40 + j
     If LenB(UFELD) = 0 Then
'      z.B. Acess: Hier noch nicht ausprogrammiert, für jedes Feld muß Länge überprüft werden
'      SET rs = cn.OpenSchema(4, Array(Empty, Empty, "& Tbl & ", Empty))  ' 4 = adSchemaColumns
'      rs.open "SELECT top 1 * FROM `" & tbl & "`" ...
     Else
     Set rs = Nothing
     myFrag rs, "SELECT * FROM `information_schema`.`columns` WHERE `table_schema` = '" & fUmwfSQL(cDB) & "' AND `table_name` = '" & fUmwfSQL(Tbl) & "' AND `column_name` = '" & fUmwfSQL(Feld) & "'", adOpenStatic, Cn, adLockReadOnly
     If Not rs.BOF Then
      Dim lassen%
      lassen = 0
'      IF j = 4 THEN Stop
      If j = 4 Then
        Dim neus2$
        FNr = 50 + j
       If IsDate(Mid$(s2(ix), 2, Len(s2(ix)) - 2)) Then
        neus2 = CDbl(CDate(Mid$(s2(ix), 2, Len(s2(ix)) - 2)))
        If neus2 < 1000 Then
         lassen = True
         s2(ix) = "'" & REPLACE$(neus2, ",", ".") & "'"
         spl2s = vNS
         For i = 0 To UBound(s2) - 1
          spl2s = spl2s & s2(i) & ","
         Next i
         spl2s = spl2s & s2(UBound(s2))
         sql = Left$(sql, p1 - 1) & spl2s & ")"
        Else
        End If
       ElseIf LenB(Mid$(s2(ix), 2, Len(s2(ix)) - 2)) = 0 Then
        FNr = 60 + j
        neus2 = "0"
        lassen = True
        s2(ix) = "'0'"
        spl2s = vNS
        For i = 0 To UBound(s2) - 1
         spl2s = spl2s & s2(i) & ","
        Next i
        spl2s = spl2s & s2(UBound(s2))
        sql = Left$(sql, p1 - 1) & spl2s & ")"
       End If
      End If
      FNr = 70 + j
      If Not lassen Then
      csql.Clear
      Dim rsdt$
      Dim neulen&
      neulen = Len(s2(ix))
      If j = 4 Then
       rsdt = "varchar"
       If rs!numeric_precision + rs!numeric_scale + 1 > neulen Then neulen = rs!numeric_precision + rs!numeric_scale + 1
      Else
       rsdt = rs!data_type
      End If
      If rsdt = "decimal" Then
       neulen = neulen + rs!numeric_scale
       csql.AppVar Array("ALTER TABLE `", Tbl, "` MODIFY COLUMN `", Feld, "` ", rsdt, "(", neulen, ",", rs!numeric_scale, ") ")
      Else
       csql.AppVar Array("ALTER TABLE `", Tbl, "` MODIFY COLUMN `", Feld, "` ", rsdt, "(", neulen, ") ")
      End If
      FNr = 90 + j
      If Not IsNull(rs!character_set_name) Then If LenB(rs!character_set_name) <> 0 Then csql.AppVar Array("CHARACTER SET ", rs!character_set_name, " ")
      If Not IsNull(rs!collation_name) Then If LenB(rs!collation_name) <> 0 Then csql.AppVar Array("COLLATE ", rs!collation_name, " ")
      If rs!is_nullable = "YES" Then csql.Append "NULL " Else csql.Append "NOT NULL "
      If IsNull(rs!column_default) Then
       If rs!is_nullable = "YES" Then csql.Append "DEFAULT NULL "
      Else
       csql.AppVar Array("DEFAULT '", fUmwfSQL(rs!column_default), "'")
      End If
      If Not IsNull(rs!column_comment) Then If LenB(rs!column_comment) <> 0 Then csql.AppVar Array("COMMENT '", fUmwfSQL(rs!column_comment), "'")
      Err.Clear
'      GoTo anfang
'      Exit Sub
' die folgenden Zeilen auskommentiert 16.6.24
'      On Error Resume Next
'      Cn.Commit ' 23.10.10: evtl. könnte auch cn.execute("SHOW VARIABLES WHERE variable_name = 'autocommit'") ausgewertet werden
'      keinetrans = (Err.Number <> 0)
'      On Error GoTo fehler
''      myEFrag "COMMIT", , Cn
      ComTrans Cn, , keinetrans
      myEFrag csql.Value, rAF, Cn, keinfehler, ErrNr, ErrDes
      If keinfehler <> 0 Then
       Ausgeb ErrDes
       syscmd 4, ErrDes
      Else
       FMeld = "Tabelle " & Tbl & ": Feld: " & Feld & " auf " & neulen & " verlängert"
       Ausgeb FMeld
       syscmd 4, FMeld
      End If ' keinfehler <> 0 Then
' Folgendes auskommentiert 16.6.24
'      Set rErr = Nothing
'      myFrag rErr, "SHOW ERRORS", adOpenStatic, Cn, adLockReadOnly
'      If Not rErr.BOF Then
'       ErrDes = rErr!Message
'       If LenB(ErrDes) <> 0 Then
'        MsgBox ErrDes
'       End If
'      End If
      End If ' not lassen
     End If ' Not rs.BOF THEN
     End If
    End If ' InStrB(sql, insAnw) <> 0 THEN
    FNr = 100 + j
    Err.Clear
'    cn.Execute sql, rAF
    GoTo anfang
    If Err.Number <> 0 Then
     MsgBox App.EXEName & ": Fehler " & Err.Number & " beim Einfügen auf Server '" & svr & "' in Datenbank " & cDB & " in Tabelle " & Tbl & ":" & vbCrLf & Err.Description
     GoTo fehler
    End If
    Exit For
   End If ' InStrB(ErrDes, dtl(j)) <> 0 THEN
   If j = 5 Then
    Debug.Print "Anderer Fehler in Inskorr: " & ErrDes
   End If
  Next j
 End If ' Err.Number <> 0 THEN
exyt:
 If obneuMode Then myEFrag "SET SESSION sql_mode='" & altmode & "'", , Cn
' IF rAF = 0 THEN
'  GoTo anfang
''  Exit Sub
' END IF
If Not keinetrans Or restarttrans Then
' If Lese.obMySQL Then
'  myEFrag "START TRANSACTION", , CN
' Else
'  CN.BeginTrans: If CN.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 1
' End If ' Lese.obMySQL Then
  BegTrans
End If ' obrestart Or restarttrans Then
Exit Sub
fehler:
    If Err.Number = 17 Then
     MsgBox App.EXEName & ": Fehler " & Err.Number & " beim Einfügen auf Server '" & svr & "' in Datenbank " & cDB & " in Tabelle " & Tbl & ":" & vbCrLf & Err.Description
    End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & "ErrDes: " & ErrDes & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in InsKorr/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' InsKorr

Public Function TabAusgeb(rEinl As ADODB.Recordset, AusgebFrm As Form, Optional obMitausgeb% = False, Optional nz$ = vbCrLf, Optional ohneKopfZ% = False, Optional SpMinÜ, Optional spmaxü, Optional mitLeerZeilen% = False, Optional AusgabeDatei$, Optional obMitZähler = 1, Optional obohneForm%, Optional Überschrift$, Optional padCaption$, Optional obappend%, Optional obOhneAufruf%, Optional mitExcel%) As CString
 Dim i%, j&, maxL%(), Zrm%(), notNum%(), F1alt, Datei$, obcsv%
 Dim TAc As New CString ' Tabausgeb für csv-Dateien
 Dim pupos&
 Dim AusgEx$, einzf$
 Dim oExcel As Object ' excel.Application
 Dim oBook As Object, oSheet As Object

' dim obPatListe%
 Const ZrmVorgabe% = 2
 On Error GoTo fehler
 pupos = InStrRev(AusgabeDatei, ".")
 If pupos = Len(AusgabeDatei) - 3 Then If LCase$(Mid$(AusgabeDatei, pupos + 1)) = "csv" Then obcsv = True
 If LenB(AusgabeDatei) <> 0 Then
  If pupos = 0 Or pupos < Len(AusgabeDatei) - 4 Or (obcsv And InStrB(AusgabeDatei, "\") = 0) Then
   Datei = pVerz & "Listen\" & AusgabeDatei & " " & Format$(Now, "dd.mm.yy hh.mm.ss") & IIf(obcsv, ".csv", ".txt")
  Else
   If InStrB(AusgabeDatei, "\") = 0 Then
    Datei = GetSpecialFolder(ssfPERSONAL) & "\" & AusgabeDatei
   Else
    Datei = AusgabeDatei
   End If
  End If ' pupos = 0 OR pupos < Len(AusgabeDatei) - 4 THEN else
 End If
 syscmd 4, "erstelle die Datei " & Datei & " ..."
 If mitExcel Then
  pupos = InStrRev(Datei, ".")
  AusgEx = Left$(Datei, pupos) & ".xls"
  'Start a new workbook in Excel
'  Set oExcel = CreateObject("Excel.Application")
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  On Error GoTo fehler
  If oExcel Is Nothing Then
   Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = False
  Set oBook = oExcel.Workbooks.Add
  oBook.SaveAs AusgEx
  Set oSheet = oBook.Worksheets(1)
 End If ' mitexcel
  
' IF rEinl.Fields.Count > 1 THEN IF LCase$(rEinl.Fields(0).Name) = "pat_id" THEN obPatListe = True
' IF obPatListe = 0 THEN
  
  Set TabAusgeb = New CString
  ReDim maxL(rEinl.Fields.COUNT)
  ReDim Zrm(rEinl.Fields.COUNT)
  ReDim notNum(rEinl.Fields.COUNT)
  For i = 0 To UBound(Zrm)
   Zrm(i) = ZrmVorgabe
  Next i
  If rEinl.State = 0 Then
   rEinl.Open
  End If
  Do While Not rEinl.EOF
   For i = 0 To rEinl.Fields.COUNT - 1
    If Len(rEinl.Fields(i).Value) > maxL(i) Then maxL(i) = Len(rEinl.Fields(i).Value)
    If Not IsNull(rEinl.Fields(i).Value) Then If Not IsNumeric(rEinl.Fields(i).Value) Then notNum(i) = True
   Next i
   rEinl.Move 1
  Loop
  If ohneKopfZ = 0 Then
   For i = 0 To rEinl.Fields.COUNT - 1
    If Len(rEinl.Fields(i).name) > maxL(i) Then maxL(i) = Len(rEinl.Fields(i).name): Zrm(i) = ZrmVorgabe - 1
   Next i
  End If
  If Not IsMissing(SpMinÜ) Then
   For i = 0 To rEinl.Fields.COUNT - 1
    If UBound(SpMinÜ) >= i Then
     If maxL(i) < SpMinÜ(i) And SpMinÜ(i) <> 0 Then
      maxL(i) = SpMinÜ(i)
     End If
    End If
   Next i
  End If
  If Not IsMissing(spmaxü) Then
   For i = 0 To rEinl.Fields.COUNT - 1
    If UBound(spmaxü) >= i Then
     If maxL(i) > spmaxü(i) And spmaxü(i) <> 0 Then
      maxL(i) = spmaxü(i)
     End If
    End If
   Next i
  End If
  TabAusgeb.Clear
  If ohneKopfZ = 0 Then
   For i = 0 To rEinl.Fields.COUNT - 1
    TabAusgeb.Append Left$(rEinl.Fields(i).name & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
    If obcsv Then TAc.AppVar Array(rEinl.Fields(i).name, ";")
'    If mitExcel Then oSheet.Range(Chr$(65 + i) & (j + 1)).Value = rEinl.Fields(i).name
     If mitExcel Then
      oSheet.Cells(j + 1, i + 1).Value = rEinl.Fields(i).name
      oSheet.Cells(j + 1, i + 1).Font.bold = True
     End If
   Next i
   TabAusgeb.Append nz
   If obcsv Then TAc.Append nz
  End If ' not ohneKopfZ
  If Not rEinl.BOF Then
   rEinl.MoveFirst
   If mitLeerZeilen Then F1alt = rEinl.Fields(0)
   Do While Not rEinl.EOF
    If mitExcel Then j = j + 1
    If mitLeerZeilen Then
     If rEinl.Fields(0) <> F1alt Then
      TabAusgeb.Append nz
      If obcsv Then TAc.Append nz
      F1alt = rEinl.Fields(0)
     End If
    End If
    For i = 0 To rEinl.Fields.COUNT - 1
     If notNum(i) Then
      TabAusgeb.Append Left$(rEinl.Fields(i).Value & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
     Else
      TabAusgeb.Append Right$(Space$(maxL(i)) & rEinl.Fields(i).Value, maxL(i)) & Space$(Zrm(i))
     End If
'     IF InStrB(rEinl.Fields(i).Value, "Becker") <> 0 THEN Stop
     If obcsv Or mitExcel Then einzf = REPLACE$(rEinl.Fields(i).Value, ";", ",")
     If obcsv Then TAc.AppVar Array(einzf, ";")
     If mitExcel Then oSheet.Cells(j + 1, i + 1).Value = einzf
    Next i
    TabAusgeb.Append nz
    If obcsv Then TAc.Append nz
    rEinl.Move 1
   Loop ' While Not rEinl.EOF
   rEinl.MoveFirst
  End If ' not rEinl.BOF
  If obMitausgeb Then AusgebFrm.Ausgeb TabAusgeb.Value, True
  If LenB(AusgabeDatei) <> 0 Then
   If obappend <> 0 Then Open Datei For Append As #317 Else Open Datei For Output As #317
   If Überschrift <> vNS Then Print #317, Überschrift
   If obcsv Then
    Print #317, TAc.Value
   Else
    Print #317, TabAusgeb.Value
   End If
   Close #317
   If obOhneAufruf = 0 Then
    zeigan Datei
   End If
  End If ' LenB(AusgabeDatei) <> 0 THEN
  If mitExcel Then
   oBook.Save
'   oExcel.Quit
   oExcel.Visible = True
  End If
  syscmd 4, "fertig mit " & Datei & "."
  If obohneForm <> 0 Then Exit Function
' ELSE ' obPatListe
' "KeinePatListe": Public-Argument für bedingte Kompilierung in den Projekteigenschaften
#If KeinePatListe = 0 Then
  Dim pad As New PatListe
  pad.PLArt = artTabAus
  Set pad.pRs = rEinl
  pad.obMitZähler = IIf(obMitZähler = 0, 0, 1)
  pad.Typisierung = AusgabeDatei
  pad.Label1 = AusgabeDatei
  pad.Label1.Left = 2000
  pad.Text1.Left = MINvb(pad.Label1.Left + MAXvb(pad.Label1.Width, Len(pad.Label1) * 80) + 50, pad.Width - 1500)
  pad.Label1.Width = pad.Width - pad.Left - 100
  Set pad.hlese = Lese
  If padCaption <> vNS Then pad.Caption = padCaption
  pad.Show
#End If
' END IF ' obPatListe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + nz + "LastDLLError: " + CStr(Err.LastDllError) + nz + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + nz + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TabAusgeb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function      ' TabAusgeb

' myFrag für Execute
Public Function myEFrag(ByRef sql$, Optional ByRef rAF&, Optional Cn As ADODB.Connection = Nothing, Optional keinfehler%, Optional ErrNr&, Optional ErrDes$, Optional gcl% = 70, Optional keinExec%) As ADODB.Recordset
 Dim rs As ADODB.Recordset
 Set myEFrag = myFrag(rs, sql, IIf(keinExec, adOpenDynamic, adOpenUnspecified), Cn, adLockReadOnly, gcl%, rAF, keinfehler, ErrNr, ErrDes)
End Function ' myEFrag

' .Execute nimmt adOpenForwardOnly, was viel schneller ist, aber nach einer Abfrage isnull(rs!Feld) rs!Feld zu null setzt
' rückwärts aufrufen: adopendynamic
' .update geht nur mit adOpenDynamic und (z.B.?) adLockOptimistic
Public Function myFrag(ByRef rs As ADODB.Recordset, ByRef sql$, _
                 Optional ByVal CursTp As ADODB.CursorTypeEnum = adOpenUnspecified, _
                 Optional ByRef Cn As ADODB.Connection = Nothing, _
                 Optional ByVal LockTp As ADODB.LockTypeEnum = adLockReadOnly, _
                 Optional ByVal gcl& = 70, _
                 Optional ByRef rAF&, _
                 Optional ByVal keinfehler%, _
                 Optional ByRef ErrNr&, _
                 Optional ByRef ErrDes$ _
                 ) As ADODB.Recordset
 Dim myru%, lauf&, cs$, ddb$
 Dim gcrs As New ADODB.Recordset
 Static fangefangen%
 Const maxru% = 2
 Dim MaxLauf&
 On Error GoTo fehler
' syscmd 4, sql
 
' If InStrB(UCase$(sql), "SELECT ZEITPUNKT") <> 0 And InStrB(UCase$(sql), "FORM") <> 0 Then Stop
 If Cn Is Nothing Then
  Set Cn = DBCn
 End If
 cs = Cn.Properties("Extended Properties")
 ddb = Cn.DefaultDatabase
' If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("show databases") ' USE `" & DefaultDatabase & "`
 If Not rs Is Nothing Then If rs.source <> "" Then If rs.source = sql Then Set myFrag = rs: Exit Function
 On Error Resume Next
 If InStr(1, sql, "GROUP_CONCAT", vbTextCompare) <> 0 Then
  For myru = 1 To maxru
  ' SELECT CHAR_LENGTH(GROUP_CONCAT(COLLATION_NAME SEPARATOR '"')) FROM INFORMATION_SCHEMA.COLLATIONS;
   Cn.Execute ("SET SESSION group_concat_max_len = " & gcl)
   ErrNr = Err.Number
   ErrDes = Err.Description
   If ErrNr = 0 Then Exit For
   Debug.Print "Myfrag: Fehler " & ErrNr & ": " & ErrDes & vbCrLf; " bei: " & sql
   DoEvents
'    Sleep 1000
'    DoEvents
   If myru = maxru - 2 Then
'     Call DBCnOpen
'     Set Cn = DBCn
      Set Cn = New ADODB.Connection
      Cn.Open cs
      Cn.DefaultDatabase = ddb
'     If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("USE `" & DefaultDatabase & "`")
    Else: On Error GoTo fehler
   End If ' myru = maxru - 2 Then
  Next myru
 End If ' InStr(1, sql, "GROUP_CONCAT", vbTextCompare) <> 0 Then
 On Error Resume Next
' Debug.Print "SQL: " & sql
 For myru = 1 To maxru
  Dim lngTime&
  lngTime = GetTickCount
  If CursTp = adOpenUnspecified And LockTp = adLockReadOnly Then
'   If InStrB(sql, "fuell") <> 0 Then Stop
   Set rs = Cn.Execute(sql, rAF)
  Else
   Set rs = New ADODB.Recordset ' If rs.State <> 0 Then rs.Close
   rs.Open sql, Cn, CursTp, LockTp
  End If ' CursTp = adOpenUnspecified And LockTp = adLockReadOnly Then
  lngTime = GetTickCount - lngTime
  If lngTime > 1000 Then
   Open pVerz & "fehler\perf.txt" For Append As #321
   Print #321, Now(), lngTime, " ms", sql
   Close #321
  End If
'  If InStrB(sql, "SELECT * FROM `dienstplan` LEFT JOIN `arten` ON `dienstplan`.artnr = `arten`.artnr") <> 0 Then Stop
  ErrNr = Err.Number
  ErrDes = "Description: " & Err.Description & " " & ", LastDllError: " & Err.LastDllError & ", Source: " & Err.source
  If ErrNr = 0 Then Set myFrag = rs: Exit Function Else
   If myru <> 1 Then ' z.B. bei gleichzeitigem BDT-Import regelmäßig auftretend
    If Not fangefangen Then
     Open pVerz & "fehler\fehler.txt" For Append As #320
     fangefangen = True
    End If
    Print #320, Now(), ErrDes, ErrNr, ": ", vbCrLf, " ", sql
    If keinfehler <> 0 Then Exit Function
    Debug.Print "MyFrag: Fehler " & Err.Number & ": " & Err.Description & vbCrLf; " bei: " & sql
    DoEvents
    Sleep 1000
    DoEvents
   End If ' myru <> 1 Then ' z.B. bei gleichzeitigem BDT-Import regelmäßig auftretend
   If myru = 1 Or myru >= maxru - 2 Then
'    Call DBCnOpen
'    Set Cn = DBCn
    Set Cn = New ADODB.Connection
    Cn.Open cs
    Cn.DefaultDatabase = ddb
 '   If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("USE `" & DefaultDatabase & "`")
   Else: On Error GoTo fehler
  End If ' ErrNr = 0 Then Set myFrag = rs: Exit Function Else
 Next myru
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
ErrDes = Err.Description
ErrNr = Err.Number
If ErrDes = "Der Vorgang ist für ein geschlossenes Objekt nicht zugelassen." Then
 If Not Cn Is Nothing Then
  If Cn.State = 0 Then
   Cn.Open
   Cn.DefaultDatabase = ddb
'   If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("USE `" & DefaultDatabase & "`")
   If Cn.State = 1 Then Resume
  End If ' Cn.State = 0 Then
 End If ' Not Cn Is Nothing Then
End If ' ErrDes = "Der Vorgang ist für ein geschlossenes Objekt nicht zugelassen." Then
If InStr(1, ErrDes, "gone away", vbTextCompare) <> 0 Then ' Or InStr(LCase$(ErrDes), "lost connection") <> 0 Then
' DBCnOpen
' Set Cn = DBCn
 Set Cn = New ADODB.Connection
 Cn.Open cs
 Cn.DefaultDatabase = ddb
 lauf = lauf + 1
 If lauf < MaxLauf Then Resume Else Resume Next
ElseIf InStr(1, ErrDes, "ANGEFORDERTEN EIGENSCHAFTEN", vbTextCompare) <> 0 Or InStr(1, ErrDes, "UNBEKANNTER FEHLER", vbTextCompare) <> 0 Then
 MaxLauf = 2
 If lauf < MaxLauf Then
  lauf = lauf + 1
'  DBCnOpen
'  Set Cn = DBCn
  Set Cn = New ADODB.Connection
  Cn.Open cs
  Cn.DefaultDatabase = ddb
'  If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("USE `" & DefaultDatabase & "`")
  Resume
 End If ' lauf < MaxLauf Then
'ElseIf InStr(UCase$(ErrDes), "LOST CONNECTION TO ") <> 0 Then
' DBCn.Close
' DBCn.Open
' Resume
ElseIf InStr(1, ErrDes, "INCORRECT", vbTextCompare) = 0 And InStr(1, ErrDes, "UNKNOWN", vbTextCompare) = 0 Then
 'Lost connection to MySQL server during query
 If ErrNr = 3704 Or ErrNr = -2147467259 Then MaxLauf = 5 Else MaxLauf = 100
 If lauf < MaxLauf Then
  lauf = lauf + 1
  If MaxLauf = 5 Then
'   DBCnOpen
'   Set Cn = DBCn
   Set Cn = New ADODB.Connection
   Cn.Open cs
   Cn.DefaultDatabase = ddb
'   If DefaultDatabase <> "" And Cn.DefaultDatabase <> DefaultDatabase Then Cn.Execute ("USE `" & DefaultDatabase & "`")
  Else
   Sleep 1000
  End If
  Resume
 End If
 If InStr(1, ErrDes, "lost connection", vbTextCompare) <> 0 Then Resume Next
End If
ErrDes = ErrDes & ": " & sql
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNr) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + ErrDes, vbAbortRetryIgnore, "Aufgefangener Fehler in myFrag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' myFrag(rs As ADODB.recorset, sql$)

'Public Function nulltest()
'Dim ru&, n&
'Dim stri$, z1 As Date, z2 As Date
'stri = "" ' "bildliche Darstellung"
'z1 = Now()
'Debug.Print z1
'For ru = 0 To 1000000000
' If stri = "" Then n = 3 ' LenB(stri) = 0 Then n = 3 => lenb ist geringfügig schneller bis doppelt so schnell
'Next ru
'z2 = Now()
'Debug.Print z2
'Debug.Print z2 - z1
'End Function



' Setzt DBCn zu (falls angegeben) CS, schließt oder löscht DBCn, öffnet es mit DBCn, zeigt nach Verbindungsstringwechsel die Fälle an
Public Function DBCnOpen(Optional cs$, Optional uid$, Optional pwd$, Optional opt&)
 Dim altCS$, lauf&
 On Error GoTo fehler
 If cs = DBCnS Then If Not DBCn Is Nothing Then If DBCn.State = 1 Then Exit Function ' 23.10.22
 If LenB(cs) <> 0 Or LenB(DBCn.ConnectionString) = 0 Then
neuverbind:
'  If Not DBCn Is Nothing Then ' ist immer erfüllt
   altCS = DBCnS ' DBCn.ConnectionString
'  End If
  If LenB(cs) = 0 Then
   cs = DBCnS ' DBCn.ConnectionString
  End If
  DBCnS = cs
  If obTrans = 1 Then Set DBCn = Nothing Else If DBCn.State <> 0 Then DBCn.Close
'  Set DBCn = Nothing ' geändert 21.10.22
  syscmd 4, "DBCnOpen " & Left$(DBCnS, InStr(DBCnS, "pwd"))
  DBCn.Open DBCnS, uid, pwd, opt
#If KeinePatListe = 0 Then
  If cs <> altCS Then
   Call fallzeig
  End If
#End If
 Else ' DBCn.ConnectionString = "" Then
  On Error Resume Next
  syscmd 4, "DBCnOpen (2) " & Left$(DBCnS, InStr(DBCnS, "pwd"))
  If DBCn.State <> 0 Then DBCn.RollbackTrans: DBCn.Close: Set DBCn = Nothing
  On Error GoTo fehler
  DBCn.Open DBCnS
End If ' DBCn.ConnectionString = "" Then
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Const laufgr% = 300
If lauf < laufgr Then
 lauf = lauf + 1
 DoEvents
 Sleep 1000
 If lauf > maxlz Then maxlz = lauf
 Resume
ElseIf lauf = laufgr Then
 If Not (LenB(cs) <> 0 Or LenB(DBCn.ConnectionString) = 0) Then
  Resume neuverbind
 End If
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DBCnOpen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DBCnOpen(CS$, uid$, pwd$, Opt&)

Public Function BegTrans(Optional Cn As ADODB.Connection = Nothing, Optional obkeinetr% = 0)
 On Error GoTo fehler
 If obkeinetr = 0 Then
  If Cn Is Nothing Then Set Cn = DBCn
  If Forms(0).obMySQL Then
'   myEFrag "START TRANSACTION", , Cn
  Else ' Lese.obMySQL Then
'   Cn.BeginTrans
  End If ' Lese.obMySQL Then
  If Err.Number = 0 Then If Cn.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 1
 End If ' obtr = 0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BegTrans/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' BegTrans(Optional CN As ADODB.Connection = DBCn)

Public Function ComTrans(Optional Cn As ADODB.Connection = Nothing, Optional obtr% = 1, Optional ByRef keinetrans%)
 On Error GoTo fehler
 If obtr = 1 Then
  If Cn Is Nothing Then Set Cn = DBCn
  If Forms(0).obMySQL Then ' Lese, MDI
'   myEFrag "COMMIT", , Cn
  Else ' Lese.obMySQL Then
'   Cn.CommitTrans
  End If ' Lese.obMySQL Then
  If keinetrans = 0 Then keinetrans = Err.Number
  If Cn.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 0
 End If ' obtr = 1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ComTrans/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ComTrans

Public Function wechsTrans(Optional Cn As ADODB.Connection = Nothing, Optional obtr% = 1)
 On Error GoTo fehler
 If obtr = 1 Then
  If Cn Is Nothing Then Set Cn = DBCn
  If Forms(0).obMySQL Then ' Lese, MDI
''   myEFrag "COMMIT", , CN ' ist bei START TRANSACTION schon dabei
  Else ' Lese.obMySQL Then
'   Cn.CommitTrans
  End If ' Lese.obMySQL Then
 End If ' obtr = 1
 If Forms(0).obMySQL Then ' Lese, MDI
'  myEFrag "START TRANSACTION", , Cn
 Else ' Lese.obMySQL Then
'  Cn.BeginTrans
 End If ' Lese.obMySQL Then
 If Cn.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in wechsTrans/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' wechsTrans

