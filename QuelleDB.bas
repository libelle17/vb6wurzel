Attribute VB_Name = "QuelleDB"
Option Explicit
Public Const CStrAcc$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public AnamneseVerZeichnis1$, StACCDB$, StFxDB$, StFtDB$, StOffDB$
'Public Const AnamneseVerZeichnis1$ = uVerz & "Anamnese\"
'Public Const StACCDB$ = AnamneseVerZeichnis1 & "quelle.mdb"
'Public Const StFxDB$ = uVerz & "FaxeinP.mdb"
'Public Const StFtDB$ = uVerz & "FotosinP.mdb"
'Public Const StOffDB$ = uVerz & "office.mdb"
Public Const Myquelle$ = "quelle"
Public Const Myquelle1$ = "quelle1"
Public Const Myquelle2$ = "quelle2"
Public Const MyFaxeinP$ = "faxeinp"
Public Const myFotosinP$ = "fotosinp"
Public Const Mykvaerzte$ = "kvaerzte"
Public Const Myoffice$ = "office"
'Public Const opti& = 2 + 4 + 8 + 32 '+ 2048 + 131072   ' 131118, 32 ' + 2048 + 16384
Public Const opti = 1 + 2 + 4 + 8 ' 32 macht die Auswahl bei PatAuswahl sehr langsam
Public Const opti1& = 8 + 32 + 2048 + 131072 + 1048576 + 2097152
Public Const opti2& = 1 + 2 + 8 + 32 + 2048 + 131072 + 1048576 + 2097152
'Public Const CSStr$ = "DRIVER={MySQL ODBC 5.1 Driver};server=" & liname & ";uid=...;pwd=...;OPTION=" & opti ' database=quelle1;
Public CStrHAE$ ' Connection String für Hausärzte
Public ifexists$, ifnotexists$, sqliif$, sqllen$, sqlALTER$, sqltodays$, sqlfloor$, sqlText$, sqlmemo$, sqlBool$, sqlDeletefrom$, sqlIGNORE$, sqlStern$, sqlPlus$, sqlAutoIncr$, sqlLong$ ' Klammer auf, Klammer zu für Tabellen- und Feldnamen in Access ``, in MySQL `` ' kla$, klz$,
'Public MyDB$ ' quelle, quelle1, quelle2
'Public ConStr$ ' Connection String
Public HAECn As New ADODB.Connection
'Public HACnS$ ' desgleichen für HACn
Public FxCn As New ADODB.Connection ' Kommentar hinter FxCn wieder entfernt 3.12.22 für FaxAkt
Public FxCnS$ ' Für Windows 7
Public FtCn As New ADODB.Connection
Public OffCn As New ADODB.Connection
Public KVÄDatei$
Public AnamneseVerZeichnis$
Public LVobMySQL As Boolean ' ob letzte Verbindung MySQL
Public FNr&

Public Enum DatenTyp
 quelleT
 HaT
 FaxT
 FotT
 OffT
End Enum ' DatenTyp
Public Enum ConDtb
' formDtb ' 0
 accDtb = 1 ' 0
 qDtb
 q1Dtb
 q2Dtb
End Enum ' ConDtb

Public Function DBCnOSchema(adSc As SchemaEnum, Crit, Optional SchemaID) As ADODB.Recordset
 On Error GoTo fehler
' SET DBCnOSchema = DBCn.OpenSchema(adSc, Crit, SchemaID)
  Set DBCnOSchema = myEFrag("SELECT * FROM information_schema.columns WHERE table_catalog='def' AND table_schema='" & IIf(DBCn.DefaultDatabase = "", Forms(0).MyDB, DBCn.DefaultDatabase) & "' AND TABLE_NAME='" & Crit(2) & "' AND is_generated='NEVER'")
' IF LenB(Spalte) <> 0 THEN
'  Do While Not EOF(DBCnOSchema)
'   IF dbcnopenschema!COLUMN_NAME = Spalte THEN Exit Do
'   DBCnOSchema.MoveNext
'  Loop
' END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DBCnOSchema/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DBCnOSchema

#If False Then
Public Function syscmd(Art%, Optional Inhalt$)
 On Error Resume Next
 Select Case Art
  Case acSysCmdSetStatus ' 4
   Forms(0).Fuß = Inhalt
  Case acSysCmdClearStatus ' 5
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
#End If

Function Zinit(obMySQL As Boolean)
 If obMySQL Then
'  kla = "`"
'  klz = "`"
  ifexists = " IF EXISTS "
  ifnotexists = " IF NOT EXISTS "
  sqliif = "IF"
  sqllen = "LENGTH"
  sqlALTER = " MODIFY"
  sqltodays = "TO_DAYS"
  sqlfloor = "FLOOR"
  sqlText = "VARCHAR"
  sqlmemo = "LONGTEXT"
  sqlBool = "TINYINT(1)"
  sqlDeletefrom = "TRUNCATE table "
  sqlIGNORE = "IGNORE "
  sqlStern = "%"
  sqlPlus = "_"
  sqlLong = "INT(10)"
  sqlAutoIncr = "INT(10) AUTO_INCREMENT"
 Else ' forms(0).obmysql
'  kla = "`"
'  klz = "`"
  ifexists = " "
  ifnotexists = " "
  sqliif = "iif"
  sqllen = "len"
  sqlALTER = " ALTER"
  sqltodays = "int"
  sqlfloor = "int"
  sqlText = "TEXT"
  sqlmemo = "memo"
  sqlBool = "bit"
  sqlDeletefrom = "DELETE FROM "
  sqlIGNORE = vNS
  sqlStern = "*"
  sqlPlus = "+"
  sqlLong = "long"
  sqlAutoIncr = "counter"
 End If ' forms(0).obmysql
End Function ' Zinit

Function UmwfSQL(q$) ' umw$(q$)
' IF LVobMySQL THEN
'  umw = replace$(replace$(Trim$(q), "\", "\\"), "'", "\'")
' Else
'  umw = replace$(replace$(Trim$(q), "\", "\\"), "'", "''")
' END IF
 UmwfSQL = doUmwfSQL(q, LVobMySQL)
End Function ' umw

' Liefert einen Connection-String, vorbehaltlich des verfügbaren Servers (LiName) im Fall von MySQL;
' öffnet die Verbindung nicht
Public Function aCStr$(DtT As DatenTyp, Optional cdTB As ConDtb, Optional AccName$, Optional DBName$, Optional obregneu%)
   Dim TBName$, Ü$
   Ü = vNS
 Static QAccDat$, KVAccDat$, FAccDat$, FotAccDat$, OffDat$
 On Error GoTo fehler
 Select Case DtT
  Case quelleT: TBName = "anamnesebogen":       Ü = "Patientendaten"
  Case HaT:     TBName = "hae":      Ü = "Hausärzte"
  Case FaxT:    TBName = "faxe":      Ü = "Faxe"
  Case FotT:    TBName = "deff":       Ü = "Fotos"
  Case OffT:    TBName = "adresse":      Ü = "Office"
 End Select
 FNr = 0
 Forms(0).dbv.obQuelle = 0
 FNr = 1
 Select Case cdTB
  Case accDtb ' 0
   If LenB(AccName) = 0 Then
    Select Case DtT
     Case quelleT ' 0
      FNr = 2
      Forms(0).dbv.obQuelle = True
      If LenB(QAccDat) = 0 Then
       On Error Resume Next
       QAccDat = Forms(0).MdB
       If LenB(QAccDat) = 0 Then QAccDat = Forms(1).MdB
       On Error GoTo fehler
       If LenB(QAccDat) = 0 Then QAccDat = StACCDB
      End If
      
      AccName = QAccDat
     Case HaT ' 1
      FNr = 3
      If LenB(KVAccDat) = 0 Then _
       KVAccDat = KVAccSuch
      AccName = KVAccDat
     Case FaxT ' 2
      FNr = 4
      If LenB(FAccDat) = 0 Then
       On Error Resume Next
       FAccDat = Forms(0).FxDB
       On Error GoTo fehler
       If LenB(FAccDat) = 0 Then FAccDat = StFxDB
      End If
      AccName = FAccDat
     Case FotT ' 2
      FNr = 5
      If LenB(FotAccDat) = 0 Then
       On Error Resume Next
       FotAccDat = Forms(0).FtDB
       On Error GoTo fehler
       If LenB(FotAccDat) = 0 Then FotAccDat = StFtDB
      End If
      AccName = FotAccDat
     Case OffT
      FNr = 6
      If LenB(OffDat) = 0 Then
       On Error Resume Next
       OffDat = Forms(0).offDB
       On Error GoTo fehler
       If LenB(OffDat) = 0 Then OffDat = StOffDB
      End If
      AccName = OffDat
    End Select
   Else ' AccName = vns THEN
    FNr = 702
    Select Case DtT
     Case quelleT: QAccDat = AccName
     Case HaT: KVAccDat = AccName
     Case FaxT: FAccDat = AccName
     Case OffT: OffDat = AccName
    End Select
   End If ' AccName = vns THEN
   FNr = 8
   aCStr = CStrAcc & AccName
   Forms(0).dbv.changeStill = True
   Forms(0).dbv.Datei = AccName
   Forms(0).dbv.changeStill = False
   On Error Resume Next ' wg. FaxDopp 25.5.08
   Forms(0).dlg.MdB = AccName
   On Error GoTo fehler
' 20.9.08
    FNr = 9
    Forms(0).dbv.ausaCStr = True
    Call Forms(0).machODBCAcc
    Forms(0).dbv.ausaCStr = False
'' folgendes neu 25.5.08 und wieder auskommentiert 20.9.08
'   Forms(0).dbv.ODBC = aCStr
' folgendes auskommentiert am 25.5.08
'   Forms(0).dbv.ODBC = CStrAcc
''   Call Forms(0).dbv.RegSpeichern
'   aCStr = Forms(0).dbv.cnVorb(vns, TBName, "Patientendaten")
    
  Case Else ' qDtb (1), q1Dtb (2), q2Dtb (3), obMySQL (-1)
      FNr = 10
   If LenB(DBName) <> 0 Then
   Else ' DBName <> vns
    Select Case DtT
     Case quelleT
      FNr = 11
      Forms(0).dbv.obQuelle = True
      Select Case cdTB
       Case True, qDtb: DBName = Myquelle
       Case q1Dtb: DBName = Myquelle1
       Case q2Dtb: DBName = Myquelle2
      End Select
     Case HaT
      FNr = 12
      DBName = Mykvaerzte
     Case FaxT
      FNr = 13
      DBName = MyFaxeinP
     Case FotT
      FNr = 14
      DBName = myFotosinP
     Case OffT
      FNr = 15
      DBName = Myoffice
    End Select
   End If ' DBName <> vns
      FNr = 16
   Forms(0).dbv.ausaCStr = True
      FNr = 17
   Call Forms(0).machODBCMy
      FNr = 18
   Forms(0).dbv.ausaCStr = False
'   dbv.Ü2 = "Benutzer"
      FNr = 19
   If DtT <> quelleT Then Forms(0).dbv.changeStill = True ' 14.9.08
      FNr = 20
'   IF DtT = quelleT THEN Forms(0).dbv.changeStill = True ' 20.9.08 Vergleich der Datenverbindung
   Forms(0).dbv.changeStill = True ' 27.3.10 aus zztest
   Forms(0).dbv.DaBa = DBName
     FNr = 21
   Forms(0).dbv.changeStill = False
'   aCStr = aCStr & TBName & ";option=" & opti
'   SET ACon = DBCnOpen((DtT), acstr)
 End Select
' folgendes hinter END SELECT geschoben am 20.9.08 aufgrund Vergleich der Datenbankstrukturen
      FNr = 22
 aCStr = Forms(0).dbv.cnVorb(DBName, TBName, Ü, obregneu, cdTB = accDtb)
      FNr = 23
 Exit Function
fehler:
ErrNumber = Err.Number
ErrLastDllError = Err.LastDllError
ErrSource = Err.source
ErrDescr = Err.Description

 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) & vbCrLf & "LastDLLError: " & CStr(ErrLastDllError) & vbCrLf & "Source: " & IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) & vbCrLf & "Description: " & ErrDescr & vbCrLf & "DtT: " & DtT & vbCrLf & "cdTB:" & cdTB & vbCrLf & "AccName: " & AccName & vbCrLf & "DBName: " & DBName & vbCrLf & "obregneu: " & obregneu, vbAbortRetryIgnore, "Aufgefangener Fehler in ACStr/" & AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aCStr()

Public Function DBNamS$(DtT As DatenTyp)
 On Error GoTo fehler
 Select Case DtT
  Case quelleT: DBNamS = Myquelle
  Case HaT: DBNamS = Mykvaerzte
  Case FaxT: DBNamS = MyFaxeinP
  Case FotT: DBNamS = myFotosinP
  Case OffT: DBNamS = Myoffice
  Case Else
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
   MsgBox "Stop in DBNamS" & vbCrLf & "DtT:" & DtT
   Stop
 End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ACStr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DBName

' in Vergleiche
Public Function acon(DtT As DatenTyp, Optional cdTB As ConDtb, Optional AccName$, Optional DBName$, Optional CnStr$, Optional obZinit%, Optional obregneu%, Optional ohneFallzeig%) As ADODB.Connection
 Static altcdTB As ConDtb
 Dim erg$
 On Error GoTo fehler
 syscmd 4, "Stelle Verbindung her (Datentyp: " & DtT & ", Tabelle: " & cdTB & ")"
 DoEvents
nochmal:
 If cdTB = 0 Then
  If Forms(0).obMySQL Then
   If DtT = quelleT Then
    If LenB(DBName) = 0 Then DBName = Forms(0).MyDB Else Forms(0).MyDB = DBName
   End If
   Select Case Forms(0).MyDB
    Case "quelle":  cdTB = qDtb
    Case "quelle1": cdTB = q1Dtb
    Case "quelle2": cdTB = q2Dtb
   End Select
  Else
   If DtT = quelleT Then If LenB(AccName) = 0 Then AccName = Forms(0).Ziel Else Forms(0).Ziel = AccName
   cdTB = accDtb
  End If
 Else
  Forms(0).obMySQL = (cdTB > accDtb) ' 25.9.08
 End If
 If cdTB <> altcdTB Then
'  forms(0).ODBC = vns
'  Dim altCS%
'  forms(0).dbv.changeStill = True
'  forms(0).dbv.CnStr = vns
'  forms(0).dbv.changeStill = altcS
 Else
  altcdTB = cdTB
 End If
 CnStr = aCStr(DtT, cdTB, AccName, DBName, obregneu)
 Select Case DtT
  Case quelleT:
   SetDBCn Forms(0).dbv.wCn, Forms(0).dbv.CnStr, ohneFallzeig
  Case HaT:
   Set HAECn = Forms(0).dbv.wCn
  Case FaxT:    Set FxCn = Forms(0).dbv.wCn: FxCnS = Forms(0).dbv.CnStr
  Case FotT:    Set FtCn = Forms(0).dbv.wCn
  Case OffT:    Set OffCn = Forms(0).dbv.wCn
 End Select
' GoTo nochmal
' END IF
' LVobMySQL = InStr(ucase$(CnStr), "MYSQL") > 0 '(Not (cDtb = accDtb))
' forms(0).obMySQL = LVobMySQL
' forms(0).obAcc = Not forms(0).obMySQL
 If Forms(0).obMySQL Then
'  obStart = True
'  forms(0).MyDB = forms(0).dbv.DaBa
'  obStart = False
  Select Case Forms(0).dbv.DaBa
   Case "quelle":  cdTB = qDtb
   Case "quelle1": cdTB = q1Dtb
   Case "quelle2": cdTB = q2Dtb
  End Select
 End If
' LVobMySQL = InStrB(UCase$(CnStr), "MYSQL") > 0 '(Not (cDtb = accDtb))
' 27.12.08: Diese Zeile ging nicht auf neu eingerichtetem Computer, da CnStr$ == "Provider=MSDASQL.1;"
 LVobMySQL = Forms(0).obMySQL
 LVobMySQL = True ' 11.10.15
 Forms(0).obMySQL = LVobMySQL
 Forms(0).obAcc = Not Forms(0).obMySQL ' Kommentar 27.12.08
 If obZinit Then Call Zinit(LVobMySQL)
 Set acon = Forms(0).dbv.wCn
' FUNCTION cnVorb$(DBName$, TBName$, Optional Ü$)
'  Call dbv.cnVorb(vns, "anamnesebogen", "Patientendaten")
'  Call forms(0).dbv.cnVorb("quelle", "anamnesebogen", "Patientendaten")
'  Call dbv.cnVorb("quelle2", "anamnesebogen", "Patientendaten")
 syscmd 5
' Forms(1).noreact = True
' Forms(1).Server = forms(0).dbv.Cpt
' Forms(1).noreact = False
 Err.Clear
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ACon/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' acon

'Public FUNCTION acon_(DtT As DatenTyp, Optional cdTB AS ConDtb, Optional AccName$, Optional DBName$, Optional CnStr$) AS ADODB.Connection
' Static Server$, runde%, CRStr$
' Dim erg$
' ON Error GoTo fehler
' syscmd 4, "Stelle Verbindung her (Datentyp: " & DtT & ", Tabelle: " & cdTB & ")"
' DoEvents
' IF cdTB = 0 THEN
'  IF Forms(0).obMySQL THEN
'   SELECT CASE Forms(0).MyDB
'    Case "quelle":  cdTB = qDtb
'    Case "quelle1": cdTB = q1Dtb
'    Case "quelle2": cdTB = q2dtb
'   END SELECT
'  Else
'   cdTB = accDtb
'  END IF
' END IF
' IF CnStr = vns OR InStrB(CnStr, DBNamS(DtT)) = 0 THEN
'  CnStr = aCStr(DtT, cdTB, AccName, DBName)
' END IF
' IF DtT = HaT AND InStrB(CnStr, "quelle") > 0 THEN Stop
'' IF instrb(ucase$(CnStr), "MYSQL") > 0 THEN
'' Else
''  IF Not Forms(0).obAcc THEN
''   Forms(0).dlg.MdB = Forms(0).dbv.Datei
''  END IF
'' END IF
' LVobMySQL = InStrB(UCase$(CnStr), "MYSQL") > 0 '(Not (cDtb = accDtb))
' Forms(0).obMySQL = LVobMySQL
' Forms(0).obAcc = Not Forms(0).obMySQL
' IF Forms(0).obMySQL THEN
'  obStart = True
'  Forms(0).MyDB = Forms(0).dbv.DaBa
'  obStart = False
'  SELECT CASE Forms(0).dbv.DaBa
'   Case "quelle":  cdTB = qDtb
'   Case "quelle1": cdTB = q1Dtb
'   Case "quelle2": cdTB = q2dtb
'  END SELECT
' END IF
' ON Error Resume Next
' IF Forms(1).Server <> vns THEN Server = Forms(1).Server
' ON Error GoTo fehler
' IF Server <> vns THEN CnStr = replace$(CnStr, LiName, Server)
' Call Zinit(LVobMySQL)
'    IF cdTB <> accDtb THEN ON Error Resume Next
'    IF runde = 0 THEN
'     CRStr = CnStr
'    Else
'     CRStr = replace$(CnStr, LiName, Server)
'    END IF
'   Do
'    SELECT CASE DtT
'     Case quelleT
'      IF Not DBCn Is Nothing THEN setdbcn vns
'      ON Error Resume Next
'      DBCnOpen CRStr
'      SET acon_ = DBCn
'     Case HaT
'      IF Not HAECn Is Nothing THEN SET HAECn = Nothing
'      HAECn.Open CRStr
'      SET acon_ = HAECn
'     Case FaxT
'      IF Not FxCn Is Nothing THEN
'       SET FxCn = Nothing
'      END IF
'      FxCn.Open CRStr
'      SET acon_ = FxCn
'     Case FotT
'      IF Not FtCn Is Nothing THEN
'       SET FtCn = Nothing
'      END IF
'      FtCn.Open CRStr
'      SET acon_ = FtCn
'     Case OffT
'      IF Not OffCn Is Nothing THEN SET OffCn = Nothing
'      OffCn.Open CRStr
'      SET acon_ = OffCn
'    END SELECT
'    Dim antw&
'    IF Err.Number = 0 OR cdTB = accDtb THEN Exit Do
'    IF Err.Number <> 0 THEN
'      antw = MsgBox("Fehlernummer: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Programm fortsetzen?", vbYesNo, App.EXEName & ": Rückfrage")
'      IF antw = vbNo THEN Ende
'      Call Forms(0).dbv.Show(1)
'      CRStr = Forms(0).dbv.CnStr
'      CnStr = Forms(0).dbv.Constr
'    END IF
'   Loop
''    IF Err.Number = -2147467259 THEN ' Datei nicht gefunden
''     MsgBox Err.Description
''     Exit Function
''    END IF
'#If False THEN
'    SELECT CASE Err.Number
'     Case ELSE '-2147467259
'      antw = MsgBox("Fehlernummer: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Anderen Datenbankserver versuchen?", vbYesNo, App.EXEName & ": Rückfrage")
'      SELECT CASE antw
'       Case vbNo: Progende
'      END SELECT
'    END SELECT
'    ON Error GoTo fehler
'    IF InStrB(UCase$(CRStr), "MYSQL") > 0 THEN
'     runde = runde + 1
'     SELECT CASE runde
'      Case 1: Server = "linmitte"
'      Case 2: Server = "mitte"
'      Case 3: Server = "linserv"
'      Case 4: Server = "server"
'      Case 5: Server = "localhost"
'      Case 6: Server = "sp0"
'      Case 7: Server = "anmeld2"
'      Case Else
'       erg = MsgBox("Fehler beim Öffnen von MySQL auf: <LiName>, linmitte, mitte, linserv, server, localhost, sp0 und anmeld2", vbRetryCancel)
'       SELECT CASE erg
'        Case vbRetry
'         runde = 0
'        Case vbCancel
'         End
'       END SELECT
'     END SELECT
'    END IF
'   Loop
'#END IF
'   syscmd 5
' ON Error Resume Next
' IF Not DBCn Is Nothing THEN
'  Forms(0).ConStri = "geöffnet: " & DBCn.ConnectionString
' Else
'  Forms(0).ConStri = "geschlossen: " & CRStr
' END IF
' IF Server = vns THEN Server = LiName
' Forms(1).noreact = True
' Forms(1).Server = Server
' Forms(1).noreact = False
' Err.Clear
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in acon_/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' acon_

'#If False THEN
'Public FUNCTION doConstrFestleg(Art AS ConDtb, obStart%, Optional MdB$, Optional frm AS Form)
' Const opti& = 2 + 4 + 8   ' 131118, 32 ' 1 + 2048 + 16384 + 131072
' SELECT CASE Art
'  Case accDtb ' 1
'   Constr = CStrAcc & MdB
''   forms(0).obMySQL = 0
'  Case Else
'   SELECT CASE Art
'    Case qDtb ' 2
'     frm.MyDB = "quelle"
'    Case q1Dtb ' 3
'     frm.MyDB = "quelle1"
'    Case q2dtb ' 4
'     frm.MyDB = "quelle2"
'   END SELECT
'   Constr$ = CStrMy & frm.MyDB & ";option=" & opti
''   forms(0).obMySQL = -1
' END SELECT
' Call Zinit(Not (Art = 0))
' IF Not obStart THEN
'' der übernächste Befehl steht im regulären Programmablauf - nicht mehr - nur hier
'  Call DBCnOpen(False, Constr, frm)
''  IF forms(0).obMySQL THEN Call myEFrag("use " & MyDB)
' END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in doConstrFestleg/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION 'doConstrFestleg(dlg AS dialog)
'#END IF

Function DatFor_k$(DaT) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 If IsNull(DaT) Then
  DatFor_k = "null"
 ElseIf (LVobMySQL) Then
  DatFor_k = "'" + Format$(DaT, "yyyy-mm-dd hh:mm:ss") + "'"
 Else
  DatFor_k = "#" + Format$(DaT, "mm\/dd\/yy hh:mm:ss") + "#"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DatFor_k/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DatFor_k

Function DatForD$(DaT) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 If IsNull(DaT) Then
  DatForD = "null"
 ElseIf (LVobMySQL) Then
  DatForD = Format$(DaT, "yyyymmdd")
 Else
  DatForD = "#" + Format$(DaT, "mm\/dd\/yy") + "#"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DatForD/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DatForD
Function SelDatum$(var$, Datum)
 Dim Beginn$
 Beginn = DatForD(Datum)
 SelDatum = " (" & var & " BETWEEN " & Beginn & " AND ADDDATE(" & Beginn & ", INTERVAL 1 DAY)) "
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SelDatum/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' SelDatum

Public Function AccOpen(ByRef accrs As ADODB.Recordset, ByRef sql$, ByRef Pfad$)
 Set accrs = Nothing
 accrs.Open sql, CStrAcc & Pfad, adOpenStatic, adLockReadOnly
End Function ' AccOpen

Public Function xlsopen(ByRef accrs As ADODB.Recordset, ByRef Pfad$)
 Const XStrb = ";Extended Properties=""Excel 8.0;HDR=no;IMEX=1"""
 Static XlsCon As New ADODB.Connection
 Static rX As New ADOX.Catalog
 Set XlsCon = Nothing
' XCon.Open XStra & Me.Datei & XStrb
 XlsCon.Open CStrAcc & Pfad & XStrb
 Set rX = Nothing
 rX.ActiveConnection = XlsCon
 accrs.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", XlsCon ' Hier Excel, nicht obmysql = 0!
End Function ' xlsopen

Public Function KVAccSuch$()
 Dim Fls As Files, Fl As File, lastDat#
 Dim FSO As New FileSystemObject
 Dim rKv As New ADODB.Recordset
 On Error GoTo fehler
   AnamneseVerZeichnis = AnamneseVerZeichnis1
   If LenB(KVÄDatei) = 0 Or Not FSO.FileExists(KVÄDatei) Then
    Call VerzPrüf(AnamneseVerZeichnis)
    Set Fls = FSO.GetFolder(AnamneseVerZeichnis).Files
    For Each Fl In Fls
     If Fl.name Like "KV*rzte*.mdb" Then
'      IF lastDat = 0 OR Fl.DateLastModified > lastDat THEN
       AccOpen rKv, "SELECT COUNT(0) ct FROM `hae`", Fl.path
       If rKv!ct > 14000 Then
        AccOpen rKv, "SELECT MAX(aktzeit) lakt FROM `hae`", Fl.path
        If rKv!lakt > lastDat Then
         KVÄDatei = Fl.path
         lastDat = rKv!lakt
        End If
       End If
'       Debug.Print Fl.Name
'       lastDat = Fl.DateLastModified
'      END IF
     End If
nächstedatei:
    Next Fl
   End If
   KVAccSuch = KVÄDatei
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147217904 Or Err.Number = -2147217865 Or Err.Number = -2147467259 Then  ' Feld nicht in Datenbank usw.
 Resume nächstedatei
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KVAccSuch/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' KVAccSuch

'Public FUNCTION KVÄVorb()
' Dim obVerändert%
' ON Error GoTo fehler
' IF (dbknr > 1) THEN
'  IF CStrHAE <> CStrMy & "kvaerzte" & ";OPTION=" & opti THEN
'   CStrHAE = CStrMy & "kvaerzte" & ";OPTION=" & opti
'   obVerändert = True
'  END IF
' Else
'  IF KVÄDatei = vns THEN
'   Call KVAccSuch
'   obVerändert = True
'  END IF
'  CStrHAE = CStrAcc & KVÄDatei
' END IF
' IF Not (dbknr > 1) AND KVÄDatei = vns THEN Call KVAccSuch
' IF obVerändert OR HAECn Is Nothing OR HAECn = vns THEN
'  IF Not HAECn Is Nothing THEN SET HAECn = Nothing
'  Call DBCnOpen(True, CStrHAE)
' END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in KVÄVorb/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' KVÄVorb

#If False Then
Public Function doSortierungÄndern()
 Dim rs As New ADODB.Recordset, rdb As ADODB.Recordset, rt As ADODB.Recordset, rc As ADODB.Recordset
 Dim rAf&, altDefDB$
    Dim sql$, sql0$
 If Not LVobMySQL Then Exit Function
 altDefDB = DefDB(DBCn)
 Set rdb = myEFrag("SHOW DATABASES WHERE `database` NOT IN ('mysql','information_schema') AND `database` = 'quelle2'")
 Do While Not rdb.EOF
  myEFrag "USE `" & rdb!Database & "`", rAf
  myEFrag ("SET foreign_key_checks=0")
'  SET rs = myEFrag("SHOW VARIABLES WHERE variable_name = 'collation_database' AND value = '" & altColl & "'")
'  IF Not rs.EOF THEN
'   myEFrag "ALTER DATABASE COLLATE '" & neuColl & "'", rAf ' "`" & rdb!Database & "`"
'  END IF
  Set rt = myEFrag("SHOW FULL TABLES WHERE table_type = 'BASE TABLE'") ' FROM `" & rdb!Database & "`
  If Not rt.BOF Then
   Do While Not rt.EOF
    sql0 = "ALTER TABLE `" & rt.Fields(0) & "`"
    sql = sql0 & " "
'    SET rs = Nothing
'    myFrag rs, "SELECT * FROM information_schema.tables WHERE table_schema = '" & rdb!Database & "' AND table_name = '" & rt.Fields(0) & "' AND table_collation = '" & altColl & "'"
'    IF Not rs.EOF THEN
'     sql = sql & " COLLATE " & neuColl & "," ' `" & rdb!Database & "`."
'     Debug.Print rAf
'    END IF
    Set rc = myEFrag("SHOW FULL COLUMNS FROM `" & rdb!Database & "`.`" & rt.Fields(0) & "` WHERE (`type` LIKE 'varchar%' OR `type` LIKE 'longtext%') AND ISNULL(`default`) AND `null`='YES'")  ' collation = '" & altColl & "'"
    Do While Not rc.EOF
     sql = sql & "MODIFY `" & rc!Field & "` " & rc!Type & " COLLATE '" & rc!collation & "' default '' NOT NULL comment '" & rc!Comment & "',"
     rc.Move 1
    Loop
    sql = Left$(sql, Len(sql) - 1)
    If sql <> sql0 Then
     myEFrag sql, rAf
'     Debug.Print rAF & ": " & sql
    End If
    rt.Move 1
   Loop ' While Not rt.EOF
   myEFrag ("SET foreign_key_checks=1")
  End If ' Not rt.BOF Then
  rdb.Move 1
 Loop ' While Not rdb.EOF
' myFrag rs, "SELECT * FROM information_schema.`COLUMNS` WHERE collation_name = 'latin1_german1_ci'"
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Ende                 ' 27.9.09 für Wien
' Exit FUNCTION ' 27.9.09 für Wien
 MsgBox "Fertig"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSortierungÄndern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'doSortierungÄndern
#End If

Public Function doSortierungÄndern0()
 Const altColl$ = "latin1_german1_ci", neuColl$ = "utf8mb4_german2_ci" 'neuColl$ = "latin1_german2_ci"
 Dim rs As New ADODB.Recordset, rdb As ADODB.Recordset, rt As ADODB.Recordset, rc As ADODB.Recordset
 Dim rAf&, altDefDB$
    Dim sql$, sql0$
 
 If Not LVobMySQL Then Exit Function
' altDefDB = DefDB(DBCn)
 Set rdb = myEFrag("SHOW DATABASES WHERE `database` NOT IN ('mysql','information_schema')")
 Do While Not rdb.EOF
  myEFrag "USE `" & rdb!Database & "`", rAf
  myEFrag ("SET foreign_key_checks=0")
  Set rs = myEFrag("SHOW VARIABLES WHERE variable_name = 'collation_database' AND value = '" & altColl & "'")
  If Not rs.EOF Then
   myEFrag "ALTER DATABASE COLLATE '" & neuColl & "'", rAf ' "`" & rdb!Database & "`"
  End If
  Set rt = myEFrag("SHOW FULL TABLES WHERE table_type = 'BASE TABLE'") ' FROM `" & rdb!Database & "`
  If Not rt.BOF Then
   Do While Not rt.EOF
    sql0 = "ALTER TABLE `" & rt.Fields(0) & "`"
    sql = sql0 & " "
    Set rs = Nothing
    myFrag rs, "SELECT 0 FROM information_schema.`tables` WHERE table_schema = '" & rdb!Database & "' AND table_name = '" & rt.Fields(0) & "' AND table_collation = '" & altColl & "'"
    If Not rs.EOF Then
     sql = sql & " COLLATE " & neuColl & "," ' `" & rdb!Database & "`."
'     Debug.Print rAF
    End If
    Set rc = myEFrag("SHOW FULL COLUMNS FROM `" & rdb!Database & "`.`" & rt.Fields(0) & "` WHERE collation = '" & altColl & "'")
    Do While Not rc.EOF
     sql = sql & "MODIFY `" & rc!Field & "` " & rc!Type & " COLLATE '" & neuColl & "' default " & IIf(IsNull(rc!Default), IIf(rc!Null = "YES", "NULL", "'' NOT NULL"), "'" & rc!Default & "' " & IIf(rc!Null = "YES", "NULL", "NOT NULL")) & " " & " comment '" & rc!Comment & "',"
     rc.Move 1
    Loop
    sql = Left$(sql, Len(sql) - 1)
    If sql <> sql0 Then
     myEFrag sql, rAf
'     Debug.Print rAF & ": " & sql
    End If
    rt.Move 1
   Loop ' While Not rt.EOF
   myEFrag ("SET foreign_key_checks=1")
  End If ' Not rt.BOF Then
  rdb.Move 1
 Loop
' myFrag rs, "SELECT * FROM information_schema.`COLUMNS` WHERE collation_name = 'latin1_german1_ci'"
'       oboffenlassen = 0    ' 27.9.09 für Wien
'       Exit FUNCTION                 ' 27.9.09 für Wien
 MsgBox "Fertig"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doSortierungÄndern0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doSortierungÄndern0


' in obAcc_Click, acon, wCnAendern, VergleichTab, fzsfuell
Public Function SetDBCn(CS As ADODB.Connection, CSStr$, Optional ohneFallzeig%)
  Dim altCS$
  On Error GoTo fehler
  If Not DBCn Is Nothing Then
   altCS = DBCnS ' DBCn.ConnectionString
  End If
  DBCnS = CSStr
'  SET DBCn = CS
  On Error Resume Next ' 19.7.23
  If obTrans Then Set DBCn = Nothing Else If DBCn.State <> 0 Then DBCn.Close
  On Error GoTo fehler
'  Set DBCn = Nothing ' geändert 21.10.22
  If DBCnS = "" Then DBCnS = altCS
  On Error Resume Next  ' 19.7.23
  DBCn.Open DBCnS
  On Error GoTo fehler
  If CS Is Nothing Then
   Exit Function
  End If
  If CS <> altCS Then
   If DBCn.State = 0 Then
    Exit Function
   End If
#If KeinePatListe = 0 Then
   If ohneFallzeig = 0 Then Call fallzeig
#End If
  End If ' CS<>altCS
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in setDBCN/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DBCnSet(CS)


' jetzt in ComputerToolsFrei

'Public FUNCTION CurDB$(DBCn AS ADODB.Connection)
' CurDB = DBCn.Properties("Current Catalog")
' IF LenB(CurDB) = 0 THEN CurDB = DefDB(DBCn)
'End Function
'
'Public FUNCTION DefDB$(DBCn AS ADODB.Connection)
' Dim spos&, sp2&, runde%, dWort$
' IF DBCn.State = 0 THEN
'  For runde = 1 To 2
'   IF runde = 1 THEN dWort = "data source=" ELSE dWort = "database="
'   spos = InStr(LCase$(DBCn), dWort)
'   IF spos <> 0 THEN
'    sp2 = InStr(spos, DBCn, ";")
'    IF sp2 = 0 THEN sp2 = Len(DBCn)
'    DefDB = Mid$(DBCn, spos + Len(dWort), sp2 - spos - Len(dWort))
'    IF (Left$(DefDB, 1) = """" AND Right$(DefDB, 1) = """") OR (Left$(DefDB, 1) = "'" AND Right$(DefDB, 1) = "'") THEN DefDB = Mid$(DefDB, 2, Len(DefDB) - 2)
'    Exit For
'   END IF
'  Next runde
' Else
'  DefDB = DBCn.DefaultDatabase
'  ON Error Resume Next
'  IF LenB(DefDB) = 0 THEN DefDB = DBCn.Properties("Data Source Name").Value
' END IF
' IF LenB(DefDB) = 0 THEN DefDB = "quelle"
'End FUNCTION ' DefDB
'

'Public FUNCTION testAktiv%()
' Dim col As New Collection, AppZahl%, i%
' ON Error GoTo fehler
' SET col = GetProcessCollection
' For i = 1 To col.Count
'  Debug.Print col(i), App.EXEName
'  IF InStrB(col(i), "|" & ProcIDFromWnd(App.ThreadID)) = 0 THEN ' nicht aktuelle Anwendung
'   IF InStrB(col(i), App.EXEName) <> 0 THEN
'    AppZahl = AppZahl + 1
'    IF AppZahl > 0 THEN
''   MsgBox "Programm wird bereits ausgeführt!"
'     testAktiv = True
'     Exit For
'    END IF
'   ElseIf InStrB(UCase$(col(i)), "NVINI") <> 0 OR InStrB(UCase$(col(i)), "NVERB") <> 0 THEN
'    testAktiv = True
'    Exit For
'   END IF
'  END IF
' Next i
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIGNORE, "Aufgefangener Fehler in testAktiv/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIGNORE: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' testAktiv

Public Function AlterBei!(bei As Date, Geb As Date)
  AlterBei = Year(bei) - Year(Geb)
  Select Case Month(Geb)
   Case Month(bei)
    Select Case Day(Geb)
     Case Is > Day(bei)
      AlterBei = AlterBei - 1
    End Select
   Case Is > Month(bei)
    AlterBei = AlterBei - 1
  End Select
End Function ' AlterBei!

Public Function ProgEnde()
 End
End Function ' ProgEnde

