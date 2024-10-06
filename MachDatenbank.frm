VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MachDatenbank 
   Caption         =   "Datenbank von aktueller Datenbank aus duplizieren / reparieren"
   ClientHeight    =   6435
   ClientLeft      =   2115
   ClientTop       =   2070
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "MachDatenbank"
   ScaleHeight     =   6435
   ScaleWidth      =   11430
   Begin VB.TextBox Port 
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton MachAlle 
      Caption         =   "&MachAlle"
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton AusgangsdatenbankWählen 
      Caption         =   "Ausgangsdatenbank w&ählen"
      Height          =   615
      Left            =   9600
      TabIndex        =   22
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton SchreibenAufCmd 
      Caption         =   "Sc&hreiben auf"
      Height          =   375
      Left            =   9600
      TabIndex        =   20
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox SchreibenAuf 
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   3840
      Width           =   6735
   End
   Begin VB.CheckBox nurSchreiben 
      Caption         =   "&nur Schreiben in:"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   3550
      Width           =   6615
   End
   Begin VB.CheckBox auchLaborX 
      Caption         =   "auch LaborX"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CheckBox auchAnamnese 
      Caption         =   "auch Anamnese"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5640
      Width           =   1695
   End
   Begin VB.OptionButton alleDaten 
      Caption         =   "alle &Daten"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6120
      Width           =   1695
   End
   Begin VB.OptionButton nurCustomizing 
      Caption         =   "nur &Customizing"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   1695
   End
   Begin VB.OptionButton keineDaten 
      Caption         =   "&keine Daten"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton DateiSuchen 
      Caption         =   "Date&i Suchen"
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Verbindungsauswahl"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
      Begin VB.CommandButton NurLauf 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox ServerZ 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton obAccess 
         Caption         =   "&Access"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton obMySQL 
         Caption         =   "M&ySQL"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox DBn 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3120
      Width           =   6615
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Start"
      Height          =   615
      Left            =   9480
      TabIndex        =   12
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CheckBox mitRelationen 
      Caption         =   "Mit &Relationen"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CheckBox mitIndices 
      Caption         =   "Mit Indi&zes"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CheckBox MitTabellen 
      Caption         =   "&Mit Tabellen (Inhalte werden erhalten)"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Ausgangsdb 
      Height          =   615
      Left            =   6240
      TabIndex        =   25
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Port_Lab 
      Caption         =   "&Port:"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Zielstring 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   10695
   End
   Begin VB.Label QuellString 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10455
   End
End
Attribute VB_Name = "MachDatenbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Comp ' Laufvariable für Computer
Dim Cpts As New Collection
Const NL0$ = "Nur Computer mit MySQL auf&listen", NL1$ = "Alle Computer auf&listen"
Dim obNurLauf%
Dim hDBn ' hiesige Defaultdatenbank
Const CptLänge% = 15

Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist
Dim cCoDB$ ' ohne Datenbank
Dim mpwd$ ' Masterpasswort
Dim cnz As New ADODB.Connection
Dim ErrNumber&, ErrDescr$, ErrLastDllError&, ErrSource$
' Dim obStart%
Enum Constr_Feld
 constraint_name
 table_name
 referenced_table_name
 referenced_column_name
 COLUMN_NAME
 update_rule
 delete_rule
End Enum
' Dim ZielVerbindung$ ' wird nicht mehr verwendet 5.10.24

Function GetServr$(DBCn As ADODB.Connection)
 Dim spos&, sp2&, cs$
 cs = DBCn.Properties("Extended Properties")
 spos = InStr(1, cs, "server=", vbTextCompare)
 If spos <> 0 Then
  sp2 = InStr(spos, cs, ";")
  If sp2 = 0 Then sp2 = Len(cs)
  GetServr = Mid$(cs, spos + 7, sp2 - spos - 7)
 End If ' spos <> 0 Then
End Function ' GetServr

' in aktualisiercon
Public Function setzmpwd$(Optional neu%)
 If neu Or mpwd = "" Then
  mpwd = InputBox("Datenbankpasswort für Benutzer `mysql`:", "Passworteingabe", mpwd)
 End If ' neu Or mpwd = "" Then
 setzmpwd = mpwd
End Function ' setzmpwd(Optional neu%)

Function setzCStrs()
 Set cnz = Nothing
 If Me.obMySQL <> 0 Then
'  cCoDB = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Trim$(LEFT(Me.ServerZ, 15)) & ";uid=mysql;pwd='" & mpwd & "';"
  Call setzmpwd
  cCoDB = "PROVIDER=MSDASQL;driver={" & ODBCStr & "};server=" & Trim$(Left$(Me.ServerZ, 15)) & ";uid=mysql;"
  cnzCStr = cCoDB & "database=" & Me.DBn & ";"
  cCoDB = cCoDB & "pwd="
  cnzCStr = cnzCStr & "pwd="
  cnz.Open cnzCStr & mpwd & ";"
 Else ' Me.obMySQL <> 0 Then
'  cnzCStr = "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=admin;Data Source=" & Me.DBn & ";Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
  cnzCStr = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Data Source=" & Me.DBn & ";Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Password="""";"
  cCoDB = cnzCStr
  cnz.Open cnzCStr
 End If ' Me.obMySQL <> 0 Then Else
End Function ' setzCStrs()

' in DateiSuchen_LostFocus, obAccess_Click, obMySQL_Click, machDatenbank.Start_Click
Function aktualisiercon%(Optional obscharf%)
 Dim oCat As New ADOX.Catalog ' SET oCat = CreateObject("ADOX.Catalog")
 Dim erg, rAf&, ErrNr&
 aktualisiercon = True
 Me.QuellString = DBCn.Properties("Extended Properties")
' Exit Function
 On Error Resume Next
 Err.Clear
 Call setzCStrs
 Lese.Ausgeb "Verbindung: " & cnzCStr & "...;" & vbCrLf & Err.Number & " " & Err.Description, 1
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", , cnz, True, ErrNr
 If ErrNr = 0 Then
'  ZielVerbindung = cnzCStr ' cnz.ConnectionString ' wird nicht mehr verwendet 5.10.24
 Else
  If Me.obMySQL <> 0 Then
   erg = vNS
  Else
   If LCase$(Right$(Me.DBn, 4)) <> ".mdb" Then Me.DBn = Me.DBn & ".mdb"
   erg = Dir(Me.DBn)
  End If
  If (LenB(erg) = 0 And obscharf) Or nurSchreiben <> 0 Then
   If nurSchreiben <> 0 Then erg = vbYes Else erg = MsgBox("Die Datenbank `" & Me.DBn & "` existiert nicht. Soll sie erstellt werden?", vbYesNo, "Rückfrage")
   If erg = vbNo Then
    aktualisiercon = False
   Else ' IF erg = vbYes THEN
    Err.Clear
    If Me.obMySQL <> 0 Then
     Set cnz = Nothing
'     cnzCStr = cCoDB & "database=" & Myquelle & ";"
'     Set cnz = Nothing
     myFrag cnz, cnzCStr & mpwd & ";", , , , , , True, ErrNumber
'     ErrNumber = Err.Number
     If ErrNumber <> 0 Then
      cnzCStr = cCoDB & mpwd & ";"
      cnz.Open cnzCStr
     End If
      Dim ergeb%
      Call sAusf("CREATE DATABASE IF NOT EXISTS `" & Me.DBn & "` CHARACTER SET utf8mb4 COLLATE utf8mb4_german2_ci;", , rAf)
      If rAf = -1 Then aktualisiercon = 0: Exit Function ' Telefonbuchsortierung Ä=AE, ß = ss; wörterbuchsortiung german1 wäre Ä=A, ß=s
'     IF rAF = 1 THEN
      Call sAusf("GRANT ALL ON " & IIf(Me.obMySQL <> 0, "`" & Me.DBn & "`.*", "DATABASE") & " TO '" & Lese.dbv.uid & "'@'%'" & IIf(Me.obMySQL <> 0, " IDENTIFIED BY '" & Lese.dbv.pwd & "' WITH GRANT OPTION", ""), , rAf)
      Call sAusf("GRANT ALL ON " & IIf(Me.obMySQL <> 0, "`" & Me.DBn & "`.*", "DATABASE") & " TO '" & Lese.dbv.uid & "'@'localhost'" & IIf(Me.obMySQL <> 0, " IDENTIFIED BY '" & Lese.dbv.pwd & "' WITH GRANT OPTION", ""), , rAf)
'     END IF
     Call sAusf("USE `" & Me.DBn & "`", , rAf)
    Else ' Me.obMySQL <> 0 Then
     oCat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & Me.DBn & ";"
    End If ' Me.obMySQL <> 0 Then Else
'    If cnz.State = 1 Then ZielVerbindung = cnzCStr ' cnz.ConnectionString ' wird nicht mehr verwendet 5.10.24
    Call setzCStrs
'    If cnz.State = 1 Then ZielVerbindung = cnzCStr ' cnz.ConnectionString ' wird nicht mehr verwendet 5.10.24
    myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", , cnz
   End If
  ElseIf Me.obMySQL = 0 Then
   Err.Clear
   cnzCStr = "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=admin;Data Source='" & Me.DBn & "';Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=True;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
   Set cnz = Nothing
   cnz.Open cnzCStr
  End If
'   IF Err.Number <> 0 AND obscharf THEN
'    erg = MsgBox("Soll " & IIf(Me.obMySQL, Myquelle, QmdB) & " geladen werden?", vbYesNo, "Rückfrage")
'    IF erg = vbYes THEN
'     Err.Clear
'     SET cnz = Nothing
'     IF Me.obMySQL THEN
'      cnz.Open "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=" & Trim$(LEFT(Me.ServerZ.List(i), 15)) & ";uid=mysql;pwd='" & mpwd & "';database=" & Myquelle & ";"
'     Else
'      cnz.Open "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=admin;Data Source=" & QmdB & ";Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
'     END IF
'    END IF
'   END IF ' erg = vns / <> "" THEN
  End If ' Err.Number <> 0 THEN
 Me.Zielstring = cnz.Properties("Extended Properties")
 Me.DBn = DefDB(cnz)
End Function ' aktualisiercon

Private Sub AusgangsdatenbankWählen_Click()
 DBVerb.Auswahl DBVerb.DaBa, vNS, "Kopierquelle auswählen"
 Me.Ausgangsdb = "Ausgangs-DB: " & IIf(DBVerb.DaBa = "", DBCn.DefaultDatabase, DBVerb.DaBa)
 Set DBCn = Nothing
 DBCnS = DBVerb.CnStr
 DBCn.Open DBCnS ' DBVerb.CnStr
 Me.QuellString = DBCn.Properties("Extended Properties")
 If Me.nurSchreiben Then Me.Zielstring = vNS
End Sub ' AusgangsdatenbankWählen_Click

Private Sub DateiSuchen_Click()
 Me.DBn = GetFileToOpen(1)
End Sub ' DateiSuchen_Click

Private Sub DateiSuchen_LostFocus()
 Call aktualisiercon
End Sub ' DateiSuchen_LostFocus

Private Sub DBn_GotFocus()
 Me.DBn.SelStart = 0
 Me.DBn.SelLength = Len(Me.DBn.Text)
End Sub ' DBn_GotFocus

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Key(KeyCode, Shift, Me)
End Sub ' Form_KeyDown

Function testDBN()
 Static accDBn$, myDBn$
 If Me.obAccess <> 0 Then
   If LenB(accDBn) = 0 Then accDBn = "c:\testQuelle.mdb"
   myDBn = Me.DBn
   Me.DBn = accDBn
 Else
   If LenB(myDBn) = 0 Then myDBn = "testDB"
   accDBn = Me.DBn
   Me.DBn = myDBn
 End If
End Function ' testDBN

Private Sub Form_Load()
' Call aktualisier
 Me.MitTabellen = 1
 Me.mitIndices = 1
 Me.mitRelationen = 1
 Me.Port = 3306
 obStart = True
' Me.obAccess = True
 Me.obMySQL = True
 obStart = False
 Me.keineDaten = True
 Me.nurSchreiben.Caption = "An&weisungen nur schreiben in:"
 Me.SchreibenAuf = App.path & "\MachDB.bas"
 Me.ServerZ = LiName
 testSchreibenAuf
 Me.Ausgangsdb = "Ausgangs-DB: " & IIf(DBVerb.DaBa = "", DBCn.DefaultDatabase, DBVerb.DaBa)
End Sub ' Form_Load

Function CheckInit()
 Me.auchAnamnese = 0
 Me.auchLaborX = 0
End Function ' CheckInit

Private Sub keineDaten_Click()
 Call CheckInit
End Sub ' keineDaten_Click

Private Sub alleDaten_Click()
 Call CheckInit
End Sub ' alleDaten_Click

Private Sub testSchreibenAuf()
 Me.SchreibenAuf.Enabled = (Me.nurSchreiben <> 0)
 Me.DBn.Enabled = (Me.nurSchreiben = 0)
End Sub ' testSchreibenAuf

Private Sub nurSchreiben_Click()
 Call testSchreibenAuf
End Sub ' nurSchreiben_Click

Private Sub obAccess_Click()
' IF Not obStart THEN
  testDBN
  Call aktualisiercon
' END IF
End Sub ' obAccess_Click

Private Sub obMySQL_Click()
 testDBN
 Call aktualisiercon
End Sub ' obMySQL_Click

Private Sub SchreibenAufCmd_Click()
 Me.SchreibenAuf = GetFileToOpen(3)
End Sub ' SchreibenAufCmd_Click

'SELECT COUNT(0) AS `ct` FROM `quelle`.`laborxwert` GROUP BY `quelle`.`laborxwert`.`RefNr`,`quelle`.`laborxwert`.`Abkü`,`quelle`.`laborxwert`.`Langname`,`quelle`.`laborxwert`.`Quelle`,`quelle`.`laborxwert`.`QSpez`,`quelle`.`laborxwert`.`AbnDat`,`quelle`.`laborxwert`.`Wert`,`quelle`.`laborxwert`.`Einheit`,`quelle`.`laborxwert`.`Grenzwerti`,`quelle`.`laborxwert`.`Kommentar`,`quelle`.`laborxwert`.`Teststatus`,`quelle`.`laborxwert`.`Erklärung`,`quelle`.`laborxwert`.`Normbereich`,`quelle`.`laborxwert`.`NormU`,`quelle`.`laborxwert`.`NormO`,`quelle`.`laborxwert`.`AuftrHinw`
Private Sub Start_Click()
 If aktualisiercon(obscharf:=True) Then
  If Not IsNull(cnz) Or Me.nurSchreiben <> 0 Then
   If cnz.State = 1 Or Me.nurSchreiben <> 0 Then
    Call dbKopier(cnz, cnzCStr, DBn.Text)
   End If ' cnz.State = 1 Or Me.nurSchreiben <> 0 Then
  End If ' Not IsNull(cnz) Or Me.nurSchreiben <> 0 Then
 End If ' aktualisiercon(obscharf:=True) Then
End Sub ' Start_Click

Private Sub MachAlle_Click()
 Dim i&, pos&, buch$
' IF aktualisier(obscharf:=True) THEN
  If Me.nurSchreiben <> 0 Then
   Do
    If LenB(Me.SchreibenAuf) = 0 Then Exit Sub
    If Not FSO.FolderExists(Me.SchreibenAuf) Then
     For i = Len(Me.SchreibenAuf) To 0 Step -1
      If Mid$(Me.SchreibenAuf, i, 1) = "\" Then
       Me.SchreibenAuf = Left$(Me.SchreibenAuf, i)
       Exit For
      End If
     Next i
    Else
     Exit Do
    End If
   Loop
   Dim altDBCn$
   altDBCn = DBCn
   If FSO.FolderExists(Me.SchreibenAuf) Then
    Dim rt As New ADODB.Recordset, altSA$
    If Right$(Me.SchreibenAuf, 1) <> "\" Then Me.SchreibenAuf = Me.SchreibenAuf & "\"
    Set rt = myEFrag("SELECT schema_name FROM information_schema.SCHEMATA S")
    altSA = Me.SchreibenAuf
    Do While Not rt.EOF
     Me.SchreibenAuf = Me.SchreibenAuf & "MachDB" & rt.Fields(0) & ".bas"
     DoEvents
     Set DBCn = Nothing
     DBCnS = lies.dbv.CoStr & "database=" & rt.Fields(0) & ";"
     DBCn.Open DBCnS
     Call dbKopier(cnz, cnzCStr, DBn.Text, , , True)
     Me.SchreibenAuf = altSA
     rt.MoveNext
    Loop
    MsgBox "Fertig mit alle Datenbankbeschreibungen fixien von Server: " & GetServr(DBCn) & vbCrLf & "auf: " & altSA
    Set DBCn = Nothing
    DBCnS = altDBCn
    DBCn.Open DBCnS
   End If
  End If
' END IF
End Sub ' MachAlle_Click

'' 5.10.24: kommt nicht vor
'Public Function doMachDB()
'' myEFrag ("CREATE DATABASE `" & DBn & "`")
'' cnz.Open Lese.dbv.Auswahl("", "anamnesebogen", "Ziel")
'#If True Then
'  cnzCStr = "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=admin;Data Source=" & QmdB & ";Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
'  Set cnz = Nothing
'  cnz.Open cnzCStr
'  DBn = "c:\fjdkalfjdk.mdb"
'#Else
'  DBn = "quelle5"
'#End If
'' DBn = InputBox("Datenbank:", "Eingabe")
'' IF Dir(DBn) <> "" THEN
''  Kill (DBn)
'' END IF
' Call dbKopier(cnz, cnzCStr, DBn)
'End Function ' doMachDB

Public Function GetFileToOpen(nr%)
 Dim Text2$
 Dim fileflags As FileOpenConstants, i
 Dim filefilter$
 On Error GoTo fehler
 'Set the text in the dialog title bar
 With CommonDialog1
 Select Case nr
  Case 1: .DialogTitle = "Erhalten-Verzeichnis": .initDir = Me.DBn
  Case 2: .DialogTitle = "Löschen-Verzeichnis": .initDir = Text2
  Case 3: .DialogTitle = "Zieldatei ": .initDir = App.path & "\" & "MachDB.bas"
 End Select
 .DialogTitle = .DialogTitle & " auswählen"
 'Set the default file name AND filter
 .Filename = vNS
 filefilter = "Verzeichnisse (*.*)|*.*|Alle Dateien (*.*)|*.*"
 .Filter = filefilter
 .FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 .flags = fileflags
 'Show the Open common dialog box
 .ShowSave
 'Return the path AND file name SELECTed or
 'Return an empty string IF the user cancels the dialog
 GetFileToOpen = .Filename
 If GetFileToOpen = "" Then Exit Function
' For i = Len(GetFileToOpen) To 0 Step -1
'  IF Mid$(GetFileToOpen, i, 1) = "\" THEN
'   GetFileToOpen = LEFT(GetFileToOpen, i)
'   Exit For
'  END IF
' Next i
 End With
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in GetFileToOpen/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetFileToOpen

' in dbKopier, dbCopyAllMyMy
Function doMachZielDatenbank(cnz As ADODB.Connection, DBCn As ADODB.Connection, DBn$, ByRef zCat As ADOX.Catalog, obZMySQL%)
 On Error GoTo fehler
 If False Then Call sAusf("DROP DATABASE " & IIf(obZMySQL, "IF EXISTS ", "") & "`" & DBn & "`;")
 If Me.nurSchreiben = 0 Then
  Call setzmpwd
  If cnz.State <> 1 Then cnz.Open cnzCStr & mpwd & ";"
 Else
  If obZMySQL <> 0 Then
'    Call Ausf(" cnzCStr = """ & replace$(cCoDB, """", """""") & """")
   Call Ausf(" IF LenB(server) = 0 THEN Server = GetServr(DbCn)")
   Call Ausf(" cnzCStr = """ & REPLACE$(cCoDB, GetServr(DBCn), """ & Server & """) & """ & mpwd & "";""")
  Else ' Me.nurSchreiben = 0 Then
'    Call Ausf(" cnzCStr = """ & replace$(ZielVerbindung, """", """""") & """")
   Call Ausf(" cnzCStr = """ & REPLACE$(cnzCStr, """", """""") & """")
  End If
  Call Ausf(" SET cnz = Nothing")
  Call Ausf(" cnz.open cnzCStr & mpwd & "";""")
 End If ' Me.nurSchreiben = 0 Then Else
 If obZMySQL Then
  Call sAusf("CREATE DATABASE IF NOT EXISTS `¤¤ & DBN & ¤¤` CHARACTER SET utf8mb4 COLLATE utf8mb4_german2_ci;") ' Telefonbuchsortierung Ä=AE, ß = ss; wörterbuchsortiung german1 wäre Ä=A, ß=s
  Call sAusf("GRANT ALL ON " & IIf(obZMySQL, "`¤¤ & DBN & ¤¤`.*", "DATABASE") & " TO '" & Lese.dbv.uid & "'@'%'" & IIf(obZMySQL, " IDENTIFIED BY '"" & pwd & ""' WITH GRANT OPTION", ""))
  Call sAusf("GRANT ALL ON " & IIf(obZMySQL, "`¤¤ & DBN & ¤¤`.*", "DATABASE") & " TO '" & Lese.dbv.uid & "'@'localhost'" & IIf(obZMySQL, " IDENTIFIED BY '"" & pwd & ""' WITH GRANT OPTION", ""))
  Call sAusf("USE `¤¤ & DBN & ¤¤`")
  Call sAusf("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ")
 ElseIf InStrB(cnzCStr, "Provider=Microsoft.Jet.OLEDB.4.0;") <> 0 Then ' cnz.ConnectionString
  If Me.nurSchreiben = 0 Then
   On Error Resume Next
   zCat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & DBn & ";"
   On Error GoTo fehler
   Set cnz = Nothing
   cnzCStr = "Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=admin;Data Source='" & Me.DBn & "';Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=True;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
   cnz.Open cnzCStr
  Else
   Call sAusf("zCat.Create ""Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=¤¤ & DBN & ¤¤;""", , , True)
   Call sAusf("cnzCStr = ""Provider=Microsoft.Jet.OLEDB.4.0;Password="""""""";User ID=admin;Data Source='¤¤ & DBN & ¤¤';Mode=Share Deny None;Extended Properties="""""""";Jet OLEDB:System database="""""""";Jet OLEDB:Registry Path="""""""";Jet OLEDB:Database Password="""""""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""""""";Jet OLEDB:Create System Database=True;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale ON Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False""")
   Call Ausf(" SET cnz = Nothing")
   Call sAusf("cnz.Open cnzCStr")
  End If
 End If ' obZMySQL / not obZMySQL
 If Me.nurSchreiben = 0 Then zCat.ActiveConnection = cnz
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doMachZielDatenbank/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doMachZielDatenbank

Function doCopyView%(qds$, QName$, cnz As ADODB.Connection, zTabName$, obQMySQL%, qCat As ADOX.Catalog)
 Dim ViewText$
 Dim ars As New ADODB.Recordset
 On Error GoTo fehler
    If obQMySQL Then
'     SET ars = Nothing
'     myFrag ars, "SELECT view_definition vd FROM information_schema.views WHERE table_schema = '" & qds & "' AND table_name LIKE '" & zTabName & "'"
     myFrag ars, "SHOW CREATE VIEW `" & qds & "`.`" & QName & "`"
     ViewText = ars("CREATE VIEW")
     ars.Close
     ViewText = Mid$(ViewText, InStr(ViewText, " AS ") + 4)
    Else
     ViewText = qCat.Views(zTabName).Command.CommandText
    End If
    On Error Resume Next
    
    Call sAusf("DROP TABLE `" & zTabName & "`", , , True)
    Call sAusf("DROP VIEW `" & zTabName & "`", , , True)
    Call sAusf("CREATE VIEW `" & zTabName & "` AS " & ViewText, , , True)
    doCopyView = 1
    On Error GoTo fehler
    ' "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & lese.dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCreateView/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCreateView(qds$, qName$, cnz AS ADODB.Connection, zTabName$, obQMySQL%)

Function doCopyAllMyMy(cnz As ADODB.Connection, qCat As ADOX.Catalog, zCat As ADOX.Catalog)
 Dim qrs As New ADODB.Recordset
 Dim TbZ&, i&, j&, zl$(), zmax&, fmax&, Str() As New CString, cts$(), ctsz&(), TName$
 Dim PrZ&
 Dim ArtZ&() ' Artzahlen (Felder,Indices,Relationen)
 Dim p1&, p2&, runde&
 On Error GoTo fehler
 
 ' Fehlen noch  Funktionen und Prozeduren
 TbZ = qCat.Tables.COUNT - 1
 If TbZ < 0 Then Exit Function
 ReDim Str(1, TbZ, 0)
 ReDim ArtZ(2, TbZ)
 ReDim ctsz(TbZ)
 For i = 0 To TbZ
  Set qrs = Nothing
  TName = qCat.Tables(i).name
  Set qrs = myEFrag("SHOW CREATE TABLE `" & TName & "`")
  Select Case qCat.Tables(i).Type
   Case "VIEW"
    Str(0, i, 0) = qCat.Tables(i).name
    Str(1, i, 0) = qrs.Fields(1)
    ctsz(i) = 1
   Case "TABLE"
    ctsz(i) = SplitNeu(qrs.Fields(1), vbLf, cts)
    If ctsz(i) > UBound(Str, 3) Then
     ReDim Preserve Str(1, TbZ, ctsz(i) + 60)
    End If
    If ctsz(i) > zmax Then zmax = ctsz(i)
    For j = 0 To ctsz(i) - 1
     Select Case Left$(cts(j), 3)
      Case "  `"
       ArtZ(0, i) = ArtZ(0, i) + 1
       Str(1, i, j) = Mid$(cts(j), 2)
      Case "  P", "  U", "  K", "  S", "  F"
       ArtZ(1, i) = ArtZ(1, i) + 1
       Str(1, i, j) = cts(j)
      Case "  C"
       ArtZ(2, i) = ArtZ(2, i) + 1
       Str(1, i, j) = cts(j)
      Case ") E"
       Str(1, i, j) = Mid$(cts(j), 2)
      Case Else
       Str(1, i, j) = cts(j)
     End Select
     If Str(1, i, j).Right(1) = "," Then Str(1, i, j).Cut (Str(1, i, j).length - 1)
     p1 = InStr(Str(1, i, j), "`")
     If p1 <> 0 Then
      p2 = InStr(p1 + 1, Str(1, i, j), "`")
      If j = 0 Then ' Tabellennamen ohne Anführungszeichen
       Str(0, i, j) = Mid$(Str(1, i, j), p1 + 1, p2 - p1 - 1)
      ElseIf j > ArtZ(0, i) + ArtZ(1, i) Then ' Constraint, hier das Symbol von "Foreign_Key" verwenden
       p1 = InStr(Str(1, i, j), "CONSTRAINT `")
       If p1 <> 0 Then
        p2 = InStr(p1, Str(1, i, j), "` FOREIGN KEY")
        Str(0, i, j) = Mid$(Str(1, i, j), p1 + 11, p2 - p1 - 11 + 1)
       End If
      Else
       Str(0, i, j) = Mid$(Str(1, i, j), p1, p2 - p1 + 1)
      End If
     End If
    Next j
   Case "LINK", "ACCESS TABLE", "SYSTEM TABLE"
   Case Else
  End Select
 Next i
 If Me.nurSchreiben <> 0 Then SchreibF1 True ' obQMySQL
 Ausf "Dim Str(1, " & TbZ & ", " & zmax & ") As New CString, ArtZ&(3, " & TbZ & ")"
 Ausf "Dim hDBn$ ' hiesiger Datenbankname"
 Ausf ""
 
 For i = 0 To TbZ
  Ausf ""
  Ausf "Sub FüllStr" & i & "()"
  For runde = 0 To 1
   If runde = 1 Then
    If ArtZ(0, i) <> 0 Then Ausf " ArtZ(0, " & i & ") = " & ArtZ(0, i)
    If ArtZ(1, i) <> 0 Then Ausf " ArtZ(1, " & i & ") = " & ArtZ(1, i)
    If ArtZ(2, i) <> 0 Then Ausf " ArtZ(2, " & i & ") = " & ArtZ(2, i)
   End If
   For j = 0 To ctsz(i) - 1
    If Str(runde, i, j).length <> 0 Then
'     Str(runde, i, j).Replace """", """"""
     Str(runde, i, j) = REPLACE(Str(runde, i, j), """", """""")
     Ausf " Str(" & runde & ", " & i & ", " & j & ") = " & """" & Str(runde, i, j) & """"
    End If
   Next j
  Next runde
  Ausf "End Sub ' FüllStr" & i
 Next i
 
 If Me.nurSchreiben <> 0 Then SchreibF2 qCat.ActiveConnection.DefaultDatabase
 
 If Me.MitTabellen <> 0 Then
  Call doMachZielDatenbank(cnz, DBCn, DBn, zCat, True) ' obZMySQL
 End If
 
 For i = 0 To TbZ
  Ausf " FüllStr" & i
 Next i
 
 On Error Resume Next
'  Call sAusf("BEGIN WORK")
 Call sAusf("SET FOREIGN_KEY_CHECKS = 0")
 On Error GoTo fehler
 
 If Me.nurSchreiben = 0 Then
  Call doGenMachDB_Direkt(TbZ, Str, ArtZ, cnz)
 Else
 Ausf ""
 Ausf " Dim j&, ZZ&, Tbl$, sql As New CString"
 Ausf " For i = 0 To " & TbZ
 Ausf "  IF InstrB(Str(1, i, 0), ""CREATE TABLE"") <> 0 THEN"
 Ausf "   Tbl = Str(0, i, 0)"
 Ausf "   ZZ = ArtZ(0, i) + ArtZ(1, i)"
 Ausf "   sql = ""CREATE TABLE IF NOT EXISTS `"" & Tbl & ""` ("" & vbLf"
 Ausf "   For j = 1 To ZZ"
 Ausf "    If Str(1, i, j) <> """" Then"
 Ausf "     sql.Append Str(1, i, j)"
 Ausf "     IF j < ZZ THEN sql.AppVar Array("","", vbLf)"
 Ausf "    End If ' If Str(1, i, j) <> """" Then"
 Ausf "   Next j"
 Ausf "   ZZ = ZZ + ArtZ(2, i) + 1"
 Ausf "   sql.AppVar Array(vbLf, "")"")"
 Ausf "   sql.Append Str(1, i, ZZ)"
 Ausf "   FNr = doEx(sql.Value, 0)"
 Ausf "   Do"
 Ausf "    SET rsc = nothing"
 Ausf "    myFrag rsc, ""SHOW CREATE TABLE `"" & tbl & ""`"", adOpenStatic, cnz, adLockReadOnly"
 Ausf "    sct = rsc.Fields(1)"
 Ausf "    IF InStrB(sct, ""CREATE ALGORITHM"") = 1 THEN"
 Ausf "     FNr = doEx(""DROP VIEW `"" & Tbl & ""`"", 0)"
 Ausf "     FNr = doEx(sql.Value, 0)"
 Ausf "    Else"
 Ausf "     Exit Do"
 Ausf "    END IF"
 Ausf "   Loop"
 Ausf "   IF InStrB(AIoZ(sct), AIoZ(Str(1, i, ZZ))) = 0 THEN"
 Ausf "    Call doEx(""ALTER TABLE `"" & tbl & ""`"" & Str(1, i, ZZ), 0)"
 Ausf "   END IF"
 Ausf "   TMt.Clear"
 Ausf "   SplitN sct, vbLf, Spli"
 Ausf "   For j = 1 To ArtZ(0, i) ' Tabellenfelder"
 Ausf "    Dim k&, enthalten%, genau%, Posi$"
 Ausf "    enthalten = 0"
 Ausf "    genau = 0"
 Ausf "    k = 0"
 Ausf "    ' SET rsc = Nothing"
 Ausf "    myFrag rsc, ""SHOW columns FROM `"" & Tbl & ""` WHERE field = '"" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & ""'"", adOpenStatic, cnz, adLockReadOnly"
 Ausf "    enthalten = Not rsc.BOF"
 Ausf "    IF enthalten THEN"
 Ausf "     genau = (InStrB(sct, Str(1, i, j)) <> 0)"
 Ausf "     IF Not genau THEN"
 Ausf "      CLen = -1 ' Column-Length nicht kürzen"
 Ausf "      obLT = (InStrB(sct, Str(0, i, j) & "" longtext"") <> 0)"
 Ausf "      IF Not obLT THEN"
 Ausf "       p1 = InStr(sct, ""("")"
 Ausf "       p2 = InStr(p1, sct, Str(0, i, j)) 'zCat.Tables(Tbl).Columns(k).Name & ""`"")"
 Ausf "       IF p2 = 0 THEN p2 = InStr(p1, LCase$(sct), LCase$(Str(0, i, j)))"
 Ausf "       p1 = InStr(p2, sct, ""("")"
 Ausf "       p3 = InStr(p2, sct, "","")"
 Ausf "       IF p3 = 0 THEN p3 = InStr(p2, sct, vbLf & "")"")"
 Ausf "       IF p1 <> 0 AND p1 < p3 THEN"
 Ausf "        p2 = InStr(p1, sct, "")"")"
 Ausf "        CLen = Mid$(sct, p1 + 1, p2 - p1 - 1)"
 Ausf "       END IF"
 Ausf "      END IF"
 Ausf "     END IF"
 Ausf "    END IF"
 Ausf "    IF Not enthalten OR Not genau THEN"
 Ausf "     IF j = 1 THEN"
 Ausf "      posi = "" FIRST,"""
 Ausf "     Else"
 Ausf "      posi = "" AFTER "" & Str(0, i, j - 1) & "","""
 Ausf "     END IF"
 Ausf "     IF Not enthalten THEN"
 Ausf "      TMt.AppVar (Array("" add "", Str(1, i, j), posi))"
 Ausf "     ElseIf Not genau THEN"
 Ausf "      IF CLen <> -1 OR obLT THEN"
 Ausf "       p1 = InStr(Str(1, i, j), ""("")"
 Ausf "       IF p1 <> 0 THEN"
 Ausf "        p2 = InStr(p1, Str(1, i, j), "")"")"
 Ausf "        IF p2 <> 0 THEN"
 Ausf "         CLen1 = Mid$(Str(1, i, j), p1 + 1, p2 - p1 - 1)"
 Ausf "         IF obLT THEN"
 Ausf "          Str(1, i, j).Replace ""varchar("" & CLen1 & "")"", ""longtext"""
 Ausf "         ElseIf CLen1 < CLen THEN"
 Ausf "          Str(1, i, j).Replace ""("" & CLen1 & "")"", ""("" & CLen & "")"""
 Ausf "         END IF ' obLT THEN ELSE"
 Ausf "         genau = (InStrB(sct, Str(1, i, j)) <> 0)"
 Ausf "        END IF ' p2 <> 0 THEN"""
 Ausf "       END IF ' p1 <> 0 THEN"""
 Ausf "      END IF ' CLen <> -1 OR obLT THEN"""
 Ausf "      IF Not genau THEN"
 Ausf "       TMt.AppVar (Array("" MODIFY "", Str(1, i, j), posi))"
 Ausf "      END IF ' Not genau THEN"
 Ausf "     END IF ' Not enthalten THEN"""
 Ausf "    END IF ' Not enthalten OR Not genau THEN"""
 Ausf "   Next j"
 Ausf "   For j = ArtZ(0, i) + 1 To ArtZ(0, i) + ArtZ(1, i) ' Indices"
 Ausf "    IF InStrB(sct, Str(1, i, j)) = 0 THEN"
 Ausf "     IF InStrB(Str(1, i, j).Value, ""PRIMARY"") <> 0 THEN"
 Ausf "      IF InStrB(sct, ""PRIMARY KEY ("") <> 0 THEN"
 Ausf "       TMt.Append ("" DROP PRIMARY KEY,"")"
 Ausf "      END IF"
 Ausf "     Else"
 Ausf "      IF InStrB(sct, ""KEY "" & Str(0, i, j).Value) <> 0 THEN"
 Ausf "       TMt.AppVar Array("" DROP KEY "", Str(0, i, j), "","")"
 Ausf "      END IF"
 Ausf "     END IF"
 Ausf "     TMt.AppVar Array("" add "", Str(1, i, j), "","")"
 Ausf "    END IF"
 Ausf "   Next j"
 Ausf "   IF TMt.Length <> 0 THEN"
 Ausf "    TMt.Cut (TMt.Length - 1)"
 Ausf "    Call doEx(""ALTER TABLE `"" & tbl & ""` "" & TMt.Value, -1)"
 Ausf "   END IF"
 Ausf "  END IF ' InStrB(Str(1, i, 0), ""CREATE TABLE"") <> 0 THEN"
 Ausf " Next i"
 
 Ausf " For i = 0 To " & TbZ
 Ausf "  IF InstrB(Str(1, i, 0), ""CREATE TABLE"")<>0 THEN"
 Ausf "   Tbl = Str(0, i, 0)"
 Ausf "   ZZ = ArtZ(0, i) + ArtZ(1, i)"
 Ausf "   ' SET rsc = nothing"
 Ausf "   myFrag rsc, ""SHOW CREATE TABLE `"" & tbl & ""`"", adOpenStatic, cnz, adLockReadOnly"
 Ausf "   sct = rsc.Fields(1)"
 Ausf "   ZZ = ZZ + ArtZ(2, i) + 1"
 
 Ausf "   For j = ArtZ(0, i) + ArtZ(1, i) + 1 To ZZ - 1 'Constraints"
 Ausf "    IF InStrB(sct, Str(1, i, j)) = 0 THEN"
 Ausf "     IF InStrB(sct, ""CONSTRAINT "" & Str(0, i, j)) <> 0 THEN"
 Ausf "      Call doEx(""ALTER TABLE `"" & Tbl & ""` DROP FOREIGN KEY "" & Str(0, i, j), 0)"
 Ausf "     END IF"
 Ausf "     Call doEx(""ALTER TABLE `"" & Tbl & ""` ADD"" & Str(1, i, j), 0)"
 Ausf "    END IF"
 Ausf "   Next j"
 Ausf "  END IF ' InStrB(Str(1, i, 0), ""CREATE TABLE"") <> 0 THEN"
 Ausf " Next i"
 Ausf " Dim runde%"
 Ausf " For runde = 0 to 5"
 Ausf "  For i = 0 To " & TbZ
 Ausf "   IF InStrB(Str(1, i, 0), ""DEFINER VIEW"") <> 0 THEN"
 Ausf "    Dim obCr%"
 Ausf "    obCr = 0"
 Ausf "    ' SET rsc = Nothing"
 Ausf "    myFrag rsc, ""SHOW TABLES FROM `"" & DBn & ""` WHERE `tables_in_"" & DBn & ""` = """""" & Str(0, i, 0) & """""""", adOpenStatic, cnz, adLockReadOnly"
 Ausf "    IF rsc.BOF THEN"
 Ausf "     obCr = True"
 Ausf "    Else"
 Ausf "     SET rsc = Nothing"
 Ausf "     myFrag rsc, ""SHOW CREATE TABLE `"" & Str(0, i, 0) & ""`"", adOpenStatic, cnz, adLockReadOnly"
 Ausf "     IF rsc.Fields(1) <> Str(1, i, 0) THEN"
 Ausf "      Call doEx(""DROP TABLE IF EXISTS `"" & Str(0, i, 0) & ""`"", 0)"
 Ausf "      Call doEx(""DROP VIEW IF EXISTS `"" & Str(0, i, 0) & ""`"", 0)"
 Ausf "      obCr = True"
 Ausf "     END IF"
 Ausf "    END IF"
 Ausf "    IF obCr THEN"
 Ausf "     Call doEx(Str(1, i, 0).Value, True)"
 Ausf "    END IF"
 Ausf "   END IF"
 Ausf "  Next i"
 PrZ = qCat.Procedures.COUNT
 Dim fspli$()
 For i = 0 To PrZ - 1
  Set qrs = Nothing
  TName = qCat.Procedures(i).name
  Dim k%
  On Error Resume Next
  For k = 0 To 20
   If k Mod 2 = 0 Then Set qrs = myEFrag("SHOW CREATE FUNCTION `" & TName & "`") Else Set qrs = myEFrag("SHOW CREATE PROCEDURE `" & TName & "`")
   If Err.Number = 0 Then Exit For
  Next k
  Const maxzz% = 20
  fspli = Split(qrs.Fields(2), Chr$(10))
  If Err.Number = 0 Then
   Ausf "  fsql = _"
   For k = 0 To UBound(fspli)
    Ausf "  """ & REPLACE$(REPLACE$(fspli(k), Chr$(13), " "), """", """""") & " """ & IIf(k = UBound(fspli) Or k Mod maxzz = 0, "", " & _")
    If k Mod maxzz = 0 And k < UBound(fspli) Then
     Ausf "  fsql = fsql & _"
    End If ' k Mod 12 = 0 And k < UBound(fspli) - 1 Then
   Next k
   Ausf "  Call doEx(fsql, True)"
  End If ' Err.Number = 0 Then
  On Error GoTo fehler
 Next i
 Ausf " Next runde"
 End If ' me.nurschreiben
 DoEvents
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCopyAllMyMy/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCopyAllMyMy

Function doEx_Direkt%(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen
 Dim rAf&, FMeld$
 Dim lErrNr&, fDesc$
 Const obProt = 0
 On Error Resume Next
 cnz.DefaultDatabase = hDBn
 If obtolerant Then On Error Resume Next Else On Error GoTo fehler
 myEFrag sql, rAf, cnz, True, lErrNr, fDesc
' lErrNr = Err.Number
' fDesc = Err.Description
 FMeld = IIf(lErrNr = 0, "Kein Fehler", "Err.Nr " & lErrNr & ", " & fDesc) & ", rAf: " & rAf & " bei " & sql
 On Error GoTo fehler
 If lErrNr <> 0 Then
  Debug.Print FMeld
  Debug.Print fDesc
 End If
 If obProt Then Print #302, FMeld
 DoEvents
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case Err.Number
 Case -2147467259 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)
  doEx_Direkt = 150
  Exit Function
End Select
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) & vbCrLf & "LastDLLError: " & CStr(Err.LastDllError) & vbCrLf & "Source: " & IIf(IsNull(Err.source), vNS, CStr(Err.source)) & vbCrLf & "Description: " & Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doex_direkt/" & AnwPfad)
 Case vbAbort: Call MsgBox(" Höre auf "): ProgEnde
 Case vbRetry: Call MsgBox(" Versuche nochmal "): Resume
 Case vbIgnore: Call MsgBox(" Setze fort "): Resume Next
End Select
End Function ' doex_direkt

Function doGenMachDB_Direkt(TbZ&, Str() As CString, ArtZ&(), cnz As ADODB.Connection)
 Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$, i&, p1&, p2&, p3&, CLen&, CLen1&
 Dim Index$()
 On Error GoTo fehler
 hDBn = cnz.DefaultDatabase
 Call doEx_Direkt("SET FOREIGN_KEY_CHECKS = 0", 0)
 Dim j&, ZZ&, Tbl$, sql As New CString
 For i = 0 To UBound(Str, 2)
  If InStrB(Str(1, i, 0), "CREATE TABLE") <> 0 Then
   Tbl = Str(0, i, 0)
   ZZ = ArtZ(0, i) + ArtZ(1, i)
   sql = "CREATE TABLE IF NOT EXISTS `" & Tbl & "` (" & vbLf
   For j = 1 To ZZ
    sql.Append Str(1, i, j)
    If j < ZZ Then sql.AppVar Array(",", vbLf)
   Next j
   ZZ = ZZ + ArtZ(2, i) + 1
   sql.AppVar Array(vbLf, ")")
   sql.Append Str(1, i, ZZ)
   FNr = doEx_Direkt(sql.Value, 0)
   Set rsc = Nothing
   myFrag rsc, "SHOW CREATE TABLE `" & Tbl & "`", adOpenStatic, cnz, adLockReadOnly
   sct = rsc.Fields(1)
   If InStrB(sct, Str(1, i, ZZ)) = 0 Then
    Call doEx_Direkt("ALTER TABLE `" & Tbl & "`" & Str(1, i, ZZ), 0)
   End If
   TMt.Clear
   SplitNeu sct, vbLf, Spli
   For j = 1 To ArtZ(0, i) ' Tabellenfelder
    Dim k&, enthalten%, genau%, Posi$
    enthalten = 0
    genau = 0
    k = 0
    Set rsc = Nothing
    myFrag rsc, "SHOW columns FROM `" & Tbl & "` WHERE field = '" & Mid$(Str(0, i, j), 2, Len(Str(0, i, j)) - 2) & "'", adOpenStatic, cnz, adLockReadOnly
    enthalten = Not rsc.BOF
    If enthalten Then
     genau = (InStrB(LCase$(sct), LCase$(Str(1, i, j))) <> 0)
     If Not genau Then
      CLen = -1 ' Column-Length nicht kürzen
      p1 = InStr(sct, "(")
      p2 = InStr(p1, LCase$(sct), IIf(Left$(Str(0, i, j), 1) = "`", vNS, "`") & LCase$(Str(0, i, j)) & IIf(Right$(Str(0, i, j), 1) = "`", vNS, "`")) 'zCat.Tables(Tbl).Columns(k).Name & "`")
      p1 = InStr(p2, sct, "(")
      p3 = InStr(p2, sct, ",")
      If p3 = 0 Then p3 = InStr(p2, sct, vbLf & ")")
      If p1 <> 0 And p1 < p3 Then
       p2 = InStr(p1, sct, ")")
       CLen = Mid$(sct, p1 + 1, p2 - p1 - 1)
      End If
     End If
    End If
    If Not enthalten Or Not genau Then
     If j = 1 Then
      Posi = " FIRST,"
     Else
      Posi = " AFTER " & Str(0, i, j - 1) & ","
     End If
     If Not enthalten Then
      TMt.AppVar (Array(" add ", Str(1, i, j), Posi))
     ElseIf Not genau Then
      If CLen <> -1 Then
       p1 = InStr(Str(1, i, j), "(")
       If p1 <> 0 Then
        p2 = InStr(p1, Str(1, i, j), ")")
        If p2 <> 0 Then
         CLen1 = Mid$(Str(1, i, j), p1 + 1, p2 - p1 - 1)
         If CLen1 < CLen Then
          Str(1, i, j).REPLACE "(" & CLen1 & ")", "(" & CLen & ")"
          genau = (InStrB(sct, Str(1, i, j)) <> 0)
         End If
        End If
       End If
      End If
      If Not genau Then
       TMt.AppVar (Array(" MODIFY ", Str(1, i, j), Posi))
      End If
     End If
    End If
   Next j
   For j = ArtZ(0, i) + 1 To ArtZ(0, i) + ArtZ(1, i) ' Indices
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(Str(1, i, j).Value, "PRIMARY") <> 0 Then
      If InStrB(sct, "PRIMARY KEY (") <> 0 Then
       TMt.Append (" DROP PRIMARY KEY,")
      End If
     Else
      If InStrB(sct, "KEY " & Str(0, i, j).Value) <> 0 Then
       TMt.AppVar Array(" DROP KEY ", Str(0, i, j), ",")
      End If
     End If
     TMt.AppVar Array(" add ", Str(1, i, j), ",")
    End If
   Next j
   If TMt.length <> 0 Then
    TMt.Cut (TMt.length - 1)
    Call doEx_Direkt("ALTER TABLE `" & Tbl & "` " & TMt.Value, -1)
   End If
   For j = ArtZ(0, i) + ArtZ(1, i) + 1 To ZZ - 1 'Constraints
    If InStrB(sct, Str(1, i, j)) = 0 Then
     If InStrB(sct, "FOREIGN KEY (" & Str(0, i, j)) <> 0 Then
      Call doEx_Direkt("ALTER TABLE `" & Tbl & "` DROP FOREIGN KEY " & Str(0, i, j), 0)
     End If
     Call doEx_Direkt("ALTER TABLE `" & Tbl & "` ADD" & Str(1, i, j), 0)
    End If
   Next j
  End If ' InStr(Str(1, i, 0), "CREATE TABLE") <> 0 THEN
 Next i
 Dim runde%
 For runde = 0 To 4
 For i = 0 To UBound(Str, 2)
  If InStr(Str(1, i, 0), "DEFINER VIEW") <> 0 Then
   Dim obCr%
   obCr = 0
   Set rsc = Nothing
   myFrag rsc, "SHOW TABLES FROM `" & DBn & "` WHERE `tables_in_" & DBn & "` = """ & Str(0, i, 0) & """", adOpenStatic, cnz, adLockReadOnly
   If rsc.BOF Then
    obCr = True
   Else
    Set rsc = Nothing
    myFrag rsc, "SHOW CREATE TABLE `" & Str(0, i, 0) & "`", adOpenStatic, cnz, adLockReadOnly
    If rsc.Fields(1) <> Str(1, i, 0) Then
     Call doEx_Direkt("DROP TABLE IF EXISTS `" & Str(0, i, 0) & "`", 0)
     Call doEx_Direkt("DROP VIEW IF EXISTS `" & Str(0, i, 0) & "`", 0)
     obCr = True
    End If
   End If
   If obCr Then
    Call doEx_Direkt(Str(1, i, 0).Value, True)
   End If
  End If
 Next i
 Next runde
 Call doEx_Direkt("SET FOREIGN_KEY_CHECKS = 1", 0)
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doGenMachDB_Direkt/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doGenMachDB_Direkt

Function doCopyTable%(td As ADOX.Table, cnz As ADODB.Connection, zTabName$, obQMySQL%, obZMySQL%, qCat As ADOX.Catalog, zCat As ADOX.Catalog, runde%, aiFName$, obai%, obTabGanzKop%)
  Dim rs As New ADODB.Recordset, ars As New ADODB.Recordset
  Dim Fldnr&, i&, j&
  Dim fld As ADODB.Field
  Dim prp As ADOX.Property
  Dim req%, jt As New CString, azl%, ucc%, altFeld$, rAf&
  Dim obauto%, Comment$
  Dim altertext As New CString, alterarr() As New CString, VStr As New CString
  Dim TabEig$
  Dim qscr$, qSpli$(), rscr As New ADODB.Recordset, IndZahl%
  obTabGanzKop = 0
  ReDim alterarr(0)
  On Error GoTo fehler
'     SET rs = Nothing
     If obZMySQL = obQMySQL Then
      Dim zname$
'      IF obZMySQL THEN ' das andere geht doch auch schnell genug
'       SET rs = call sausf("SHOW TABLES FROM `" & DefDB(cnZ) & "` WHERE tables_in_" & DefDB(cnZ) & " = '" & zTabName & "'")
'       IF rs.BOF THEN zname = "" ELSE zname = rs.Fields(0)
'      Else
      zname = vNS
      On Error Resume Next
      zname = zCat.Tables(zTabName).name
      On Error GoTo fehler
'     END IF
      If obZMySQL And Me.nurSchreiben = 0 Then
       If LCase$(cnz.Properties("Server Name")) = LCase$(DBCn.Properties("Server Name")) Then ' nicht MySQL von verschiedenen Rechnern
        If LenB(zname) = 0 Then
         Call sAusf("CREATE TABLE `" & zTabName & "` LIKE `" & DefDB(DBCn) & "`.`" & td.name & "`", , , 0)
         obTabGanzKop = True
         doCopyTable = 1
         Exit Function
        End If ' LenB(zname) = 0 THEN
       End If ' LCase$(cnz.Properties("Server Name")) =
       On Error GoTo fehler
      End If ' obZMySQL AND Me.nurSchreiben = 0 THEN
     End If ' obZMySQL = obQMySQL THEN
     If runde = 1 Then
      If obQMySQL Then
       myFrag rscr, "SHOW CREATE TABLE `" & td.name & "`"
       qscr = rscr.Fields(1)
       TabEig = Mid$(qscr, InStr(qscr, " ENGINE="))
       Dim p1&, p2&
       p1 = InStr(TabEig, "AUTO_INCREMENT=")
       If p1 <> 0 Then
        p2 = InStr(p1, TabEig, " ")
        If p2 = 0 Then p2 = Len(TabEig)
        TabEig = Left$(TabEig, p1 - 1) & Right$(TabEig, Len(TabEig) - p2)
       End If
      Else
       TabEig = " ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_german2_ci"
      End If
     End If
     If LenB(zname) = 0 Or Me.nurSchreiben <> 0 Then
      If obZMySQL = 0 Then On Error Resume Next
      If runde = 1 Then
       Call Ausf(" TabEig=""" & TabEig & """")
       Call sAusf("CREATE TABLE " & IIf(obZMySQL, "IF NOT EXISTS ", "") & "`" & zTabName & "`" & IIf(obZMySQL, " (dummerl char(0))"" & TabEig & "" ", ""), , , 0)
       If obZMySQL Then
'        Call Ausf(" SET rsc = nothing")
        Call Ausf(" myFrag rsc, ""SHOW CREATE TABLE `" & zTabName & "`"", adOpenStatic, cnz, adLockReadOnly")
        Call Ausf(" sct = rsc.Fields(1)")
        Call Ausf(" IF InStrB(sct, TabEig) = 0 THEN")
       End If
       Call sAusf("ALTER TABLE `" & zTabName & "`""" & IIf(obZMySQL, "  & TabEig & "" ", ""), , , 0, IIf(obZMySQL, 2, 1))
       If obZMySQL Then
        Call Ausf(" END IF")
       End If
      End If ' runde = 1
      If Not obZMySQL Then On Error GoTo fehler
     End If ' LenB(zname) = 0 OR Me.nurSchreiben <> 0 THEN
     aiFName = vNS
     obai = 0
     altertext.Clear
     If obQMySQL Then
      myFrag rs, "SELECT * FROM `" & td.name & "` LIMIT 1"
     Else
      myFrag rs, "SELECT top 1 * FROM `" & td.name & "`"
     End If
     ReDim alterarr(rs.Fields.COUNT - 1)
     Fldnr = 0
     Lese.Ausgeb vbCrLf, False
     If obZMySQL And runde = 1 Then
      VStr = " IF "
      
     End If
      For Each fld In rs.Fields
       Lese.Ausgeb fld.name & "/" & zTabName & "...", False
       obauto = 0
       If obQMySQL Then
        If fld.Properties.COUNT > 0 Then
         If fld.Properties("isautoincrement") Then
          obauto = True
          obai = True
          aiFName = fld.name
         End If
        End If
       Else ' obQMySQL => not obQMySQL
        If qCat.Tables(td.name).Columns(fld.name).Properties("autoincrement") Then
         obauto = True
         obai = True
         aiFName = fld.name
        End If
       End If '  ' obQMySQL / not obQMySQL
'      SELECT CASE fld.Name
'       Case "Beinbefund", "Hyperkeratosen", "RR", "BeinödVen", "Weitere Befund"
'       Case Else
'      END SELECT
       Comment = vNS
       If obQMySQL Then
        If runde = 1 Or (runde = 2 And Not obZMySQL) Then
         Set ars = Nothing
         myFrag ars, "SHOW FULL COLUMNS FROM `" & td.name & "` WHERE field = '" & fld.name & "'"
        End If
       End If
       If (runde = 1 And obZMySQL) Or (runde = 2 And Not obZMySQL) Then
        If obQMySQL Then
         Comment = ars!Comment
         If runde = 2 Then ars.Close
        Else
         Comment = qCat.Tables(td.name).Columns(fld.name).Properties("description").Value
        End If ' obQMySQL
       End If ' (runde = 1 AND obZMySQL) OR (runde = 2 AND NOT obZMySQL) THEN
       If runde = 1 Then
        If obQMySQL Then
         If ars!Null = "YES" Then req = False Else req = True
        Else
         If qCat.Tables(td.name).Columns(fld.name).Attributes And adColNullable Then req = False Else req = True
        End If
        If obZMySQL Then
'        IF fld.Name = "Größe" OR fld.Name = "Gewicht" OR fld.Name = "NurDatum" OR fld.Name = "NurZeit" THEN
         Debug.Print fld.name
         VStr.AppVar (Array("InStrB(sct, ""`", fld.name, "`"") = 0 AND "))
'         IF LCase$(fld.Name) = "zp1" THEN
         jt = MySqlTyp(fld.Type, fld.DefinedSize, obauto, Not obQMySQL)
         jt.LCase
         If jt.Instr("not null auto_increment key") <> 0 Then
          jt.REPLACE "not null auto_increment key", "NOT NULL AUTO_INCREMENT KEY"
         End If
        Else ' obZMySQL => not obZMySQL
'        IF LCase$(fld.Name) = "prim" THEN ...
         jt = JetTyp(fld.Type, fld.DefinedSize, obauto)
         If jt.Instr("TEXT") <> 0 Or jt.Instr("CHAR") <> 0 Then
          jt.Append " WITH COMPRESSION"
         End If
        End If ' ' obZMySQL / not obZMySQL
        If obZMySQL Then
         If obQMySQL Then
          If Not IsNull(ars!collation) Then
           jt.AppVar Array(" COLLATE ", ars!collation)
          End If
          ars.Close
         End If
        End If ' obzmysql
        If jt.Instr("NULL") = 0 Then
         If req Then
          jt.Append " NOT NULL"
         Else ' InStrB(jt, "NULL") = 0 => <> 0
          If obZMySQL Then
           jt.Append " DEFAULT"
          End If
          jt.Append " NULL"
         End If
        End If ' InStrB(jt, "NULL") = 0 / <> 0
        If obZMySQL Then
         jt.AppVar Array(" COMMENT '", REPLACE$(Comment, "'", "''"), "'")
         If altFeld = "" Then
          jt.Append " FIRST"
         Else
          jt.AppVar Array(" AFTER `", altFeld, "`")
         End If
        End If
        Dim altname$
        altname = vNS
        On Error Resume Next
        altname = zCat.Tables(zTabName).Columns(fld.name).name
        On Error GoTo fehler
        Dim fallsDoppelt%
        fallsDoppelt = True
        If obZMySQL Then
         alterarr(Fldnr).Clear
'         IF nurSchreiben = 0 OR LenB(altname) = 0 THEN
          alterarr(Fldnr).Append " ADD `"
'         Else
'          alterarr(fldnr).Append " MODIFY `"
'         END IF
         alterarr(Fldnr).AppVar Array(fld.name, "` ")
         If nurSchreiben <> 0 Or LenB(altname) = 0 Then
          alterarr(Fldnr).Append jt
         Else
          jt.REPLACE "AUTO_INCREMENT KEY", "AUTO_INCREMENT"
          alterarr(Fldnr).Append jt
         End If
         altertext.Append alterarr(Fldnr)
         altertext.Append ","
         Fldnr = Fldnr + 1
        Else ' obZMySQL
         If LenB(altname) = 0 Then
          Call sAusf("ALTER TABLE `" & zTabName & "` add COLUMN `" & fld.name & "` " & jt, , , 0)
         Else ' LenB(altname) = 0 => <> 0
          On Error Resume Next
          Err.Clear
          Call sAusf("ALTER TABLE `" & zTabName & "` alter COLUMN `" & fld.name & "` " & jt, , , 0)
          If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description
          On Error GoTo fehler
         End If ' LenB(altname) = 0 THEN
        End If ' obZMySQL
        fallsDoppelt = False
       Else ' runde = 1 => 2
        If Not obZMySQL Then
         If Me.nurSchreiben = 0 Then
          zCat.Tables(zTabName).Columns(fld.name).Properties("Description").Value = Comment
         End If
         If Not obQMySQL Then
          For Each prp In qCat.Tables(td.name).Columns(fld.name).Properties
'         Debug.Print qCat.Tables(td.name).Columns(fld.name).Properties(prp.name).name, ":", qCat.Tables(td.name).Columns(fld.name).Properties(prp.name).Value
           On Error Resume Next
           zCat.Tables(zTabName).Columns(fld.name).Properties(prp.name).Value = prp.Value
           On Error GoTo fehler
          Next prp
         End If ' obQMySQL THEN
        End If ' not obZMySQl
        If obQMySQL Then
         azl = True
         ucc = True
        Else
         azl = qCat.Tables(td.name).Columns(fld.name).Properties("Jet OLEDB:Allow Zero Length")
         ucc = qCat.Tables(td.name).Columns(fld.name).Properties("Jet OLEDB:Compressed UNICODE Strings")
        End If
        If obZMySQL Then
        Else
         If Me.nurSchreiben = 0 Then
          Select Case zCat.Tables(zTabName).Columns(fld.name).Type
           Case 129, 130, 200, 202
            zCat.Tables(zTabName).Columns(fld.name).Properties("Jet OLEDB:Allow Zero Length") = azl ' AllowZeroLength
           Case Else
          End Select
         End If ' me.nurschreiben
         On Error Resume Next
         Err.Clear
         zCat.Tables(zTabName).Columns(fld.name).Properties("Jet OLEDB:Compressed UNICODE Strings") = ucc
         If Err.Number = 0 Then
          MsgBox "Stop in doCopyTable durch err.number = 0 bei 'zCat.Tables(zTabName).Columns(fld.Name).Properties('Jet OLEDB:Compressed UNICODE Strings') = ucc' in doCopyTable"
          Stop
         End If
         On Error GoTo fehler
        End If
       End If
       altFeld = fld.name
       Lese.Ausgeb ".", False
      Next fld

'         Dim fldnr&
         If runde = 1 And obZMySQL Then
          Lese.Ausgeb "Schreibe Tabelle `" & zTabName & "` ...", False
          On Error Resume Next
          Err.Clear
'          altertext.Cut altertext.Length - 1
'          vStr.Cut (vStr.Length - 4)
'          vStr.Append " THEN"
'          Call Ausf(vStr.Value)
'          Call sAusf("ALTER TABLE `" & zTabName & "` " & altertext, , , True, 1)
'          Call Ausf(" END IF")
          If Err.Number <> 0 Or nurSchreiben <> 0 Then
           On Error GoTo fehler
'           IF nurSchreiben <> 0 THEN Ausf " IF lErrNr <> 0 THEN"
           Call Ausf(" TMt.Clear")
           Call Ausf(" SplitN sct,vblf,Spli")
           For Fldnr = 0 To UBound(alterarr)
            Call Ausf(" tStr = """ & Mid$(alterarr(Fldnr), 5) & """")
            Call Ausf(" IF InStrB(sct, "" `" & rs.Fields(Fldnr).name & "`"") = 0 THEN")
'            Call sAusf("ALTER TABLE `" & zTabName & "` add "" & tStr & "" ", , , True, 1)
            Call Ausf("  TMt.AppVar(Array("" add "", tStr ,"",""))")
            Call Ausf(" ELSE ' IF InStrB(")
            Call Ausf("  IF InStrB(tstr, Mid$(Spli(" & Fldnr + 1 & "),2,len(Spli(" & Fldnr + 1 & "))-2)) = 0 THEN")
'            Call sAusf("ALTER TABLE `" & zTabName & "` MODIFY "" & tStr & "" ", , , True, 2)
            Call Ausf("   TMt.appvar(array("" MODIFY "", tStr ,"",""))")
            Call Ausf("  END IF")
            Call Ausf(" END IF")
           Next Fldnr
'           IF nurSchreiben <> 0 THEN Ausf " END IF ' IF lErrNr <> 0 THEN"
          End If
'          IF nurSchreiben <> 0 THEN
'           altertext.Replace " add `", " MODIFY `"
'           altertext.Replace "AUTO_INCREMENT KEY", "AUTO_INCREMENT"
'           Call sAusf("ALTER TABLE `" & zTabName & "` " & altertext)
'          END IF
          If obZMySQL Then
           Call Ausf(" IF InStrB(sct, "" `dummerl`"") <> 0 THEN")
'           Call sAusf("ALTER TABLE `" & zTabName & "` drop column dummerl", , , True, 1)
           Call Ausf("  TMt.Append("" drop COLUMN `dummerl`,"")")
           Call Ausf(" END IF")
          End If ' obZMySQL THEN
          If obZMySQL And obQMySQL And Me.nurSchreiben <> 0 And Me.mitIndices Then
           SplitNeu qscr, vbLf, qSpli
           IndZahl = 0
           For i = UBound(alterarr) + 2 To UBound(qSpli)
            Select Case Mid$(qSpli(i), 3, 3)
             Case "PRI", "UNI", "KEY", "SPA", "FUL"
              IndZahl = i
             Case Else
              Exit For
            End Select
           Next i
           Call Ausf(" redim Index(" & IndZahl - UBound(alterarr) - 2 & ")")
           For i = UBound(alterarr) + 2 To IndZahl
            Call Ausf(" Index(" & i - UBound(alterarr) - 2 & ") = """ & IIf(Right$(qSpli(i), 1) = ",", Left$(qSpli(i), Len(qSpli(i)) - 1), qSpli(i)) & """")
           Next i
           Call Ausf(" For i = 0 to " & IndZahl - UBound(alterarr) - 2)
'           Call Ausf("  SELECT CASE MID(Spli(i),3,3)")
'           Call Ausf("   case ""PRI"",""UNI"",""KEY"",""SPA"",""FUL"" ")
'           Call Ausf("    tStr = left$(spli(i),len(spli(i))-1)")
           Call Ausf("   IF instrb(sct,Index(i)) = 0 THEN")
           Call Ausf("    IF instrb(Index(i),""PRIMARY"")<> 0 THEN")
           Call Ausf("     tmt.append("" DROP PRIMARY KEY,"")")
           Call Ausf("    else")
           Call Ausf("     p1 = instr(Index(i),""`"")")
           Call Ausf("     p2 = instr(p1+1,Index(i),""`"")")
           Call Ausf("     tStr = MID(Index(i),p1,p2-p1+1)")
           Call Ausf("     IF instrb(sct,tStr)<> 0 THEN")
           Call Ausf("      TMt.Appvar(Array("" drop key "", tStr,"",""))")
           Call Ausf("     END IF")
           Call Ausf("    END IF")
           Call Ausf("   tmt.appvar(array("" add "",Index(i), "",""))")
           Call Ausf("   END IF")
'           Call Ausf("   case else")
'           Call Ausf("    exit for")
'           Call Ausf("  END SELECT")
           Call Ausf(" next i")
          End If
          Call Ausf(" IF TMt.Length <> 0 THEN")
          Call Ausf("  TMt.cut(TMt.Length-1)")
          Call sAusf("ALTER TABLE `" & zTabName & "` "" & TMt.Value & "" ", , , True, 1)
          Call Ausf(" END IF")
          If nurSchreiben <> 0 Then
           Call Ausf(" Redim Spli(0)")
           Call Ausf("")
          End If
          On Error GoTo fehler
         End If ' runde = 1 AND obZMySQL THEN
         doCopyTable = 1
 Exit Function
fehler:
 If Err.Number = -2147217887 And fallsDoppelt Then Resume Next
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCopyTable/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCopyTable

Function doCopyIndices(td As ADOX.Table, cnz As ADODB.Connection, zTabName$, obQMySQL%, obZMySQL%, zCat As ADOX.Catalog, aiFName$, obai%, qds$, zds$)
   Dim ars As New ADODB.Recordset, arsz As New ADODB.Recordset
   Dim i&, sql$
   Dim ind As ADOX.Index
   Dim obPSZ% ' = ob ars Primärschlüssel ist
   Dim col As ADOX.Column
   Dim zwi$
   
   On Error GoTo fehler
   If obZMySQL Then sql = vNS
   If obQMySQL Then
    Set ars = Nothing
    myFrag ars, "SHOW index FROM `" & qds & "`.`" & td.name & "`"
    If Not ars.BOF Then
     Do
      Dim zählen%
      zählen = 0
      If ars.EOF Then zählen = True Else If ars!seq_in_index = 1 Then zählen = True
      If zählen Then
       If LenB(sql) <> 0 Then
        sql = Left$(sql, Len(sql) - 1) & ")"
        If Not obZMySQL Then
'         call sausf ("ALTER TABLE `" & zTabName & "` " & sql)
         Call sAusf(sql & IIf(obPSZ, " WITH PRIMARY", ""))
         Lese.Ausgeb " Index für `" & zTabName & "`:" & sql & " erstellt", True
         sql = vNS
        Else
         sql = sql & ","
        End If
       End If
       If ars.EOF Then Exit Do
'       IF obai AND (aiFName = ars!column_name) AND (obZMySQL AND (ars!key_name = "PRIMARY")) OR (Not obZMySQL AND zcat.Tables(td.Name).Indexes(ars!key_name).PrimaryKey) THEN
'        obai = False
'        aiFName = vns
       Dim obsngibt%
       Dim obPSZuNE% ' = ob Primärschlüssel in der Quelle und nicht oben erstellt
'      => dann nicht löschen und nur erstellen, falls im Ziel nicht vorhanden
       obPSZ = False
       If ars!key_name = "PRIMARY" Then obPSZ = True
       obsngibt = False
       If Me.nurSchreiben <> 0 Then
        If obPSZ Then obsngibt = True
       Else ' Me.nurSchreiben <> 0 THEN
        If obZMySQL Then
         Set arsz = Nothing
         On Error Resume Next
         myFrag arsz, "SHOW index FROM `" & zds & "`.`" & zTabName & "` WHERE key_name = '" & ars!key_name & "'", adOpenStatic, cnz, adLockReadOnly ' AND column_name = '" & ars!column_namr & "'"
         If arsz.State = 1 Then If Not arsz.EOF Then obsngibt = True
         On Error GoTo fehler
         Set arsz = Nothing
        Else
         For i = 0 To zCat.Tables(td.name).Indexes.COUNT - 1
          If zCat.Tables(td.name).Indexes(i).name = ars!key_name Then
           obsngibt = True
'          IF zcat.Tables(td.Name).Indexes(i).PrimaryKey THEN obPSZ = True
           Exit For
          End If
         Next i
        End If
       End If ' Me.nurSchreiben <> 0 THEN
       obPSZuNE = obPSZ
'       IF obPSZuNE THEN IF obai AND (aiFName = ars!column_name) THEN obPSZuNE = False ' Kommentar 11.10.08
       If obsngibt And Not obPSZuNE Then
         sql = sql & "DROP " & schl(ars!key_name, obZMySQL)
         If Not obZMySQL Then
          sql = sql & " ON `" & zTabName & "`"
          Call sAusf(sql)
          sql = vNS
         Else
          sql = sql & ","
         End If
       End If
       If Not obsngibt Or (obsngibt And Not obPSZuNE) Then
        If obZMySQL Then
         sql = sql & "ADD " & IIf(ars!non_unique = 1 Or obPSZ, vNS, "UNIQUE ") & schl(ars!key_name, obZMySQL) & "("
        Else
         sql = "CREATE " & IIf(ars!non_unique = 1, vNS, "UNIQUE ") & "INDEX `" & ars!key_name & "` ON `" & zTabName & "` ("
        End If
       End If
      End If ' Zählen
      If Not obsngibt Or (obsngibt And Not obPSZuNE) Then
       sql = sql & "`" & ars!COLUMN_NAME & "`"
       If Not IsNull(ars!sub_part) Then
        sql = sql & "(" & ars!sub_part & ")"
       End If
'       zwi = ars!column_name
'       IF td.Columns(zwi).Type = adLongVarBinary OR td.Columns(zwi).Type = adLongVarChar OR td.Columns(zwi).Type = adLongVarWChar THEN
'        IF obZMySQL THEN sql = sql & "(20)"
'       END IF
       sql = sql & IIf(ars!collation <> "A", " DESC", vNS) & ","
      End If
      ars.Move 1
     Loop
     If obZMySQL And LenB(sql) <> 0 Then
      Call sAusf("ALTER TABLE `" & zTabName & "` " & Left$(sql, Len(sql) - 1), , , True)
      Lese.Ausgeb " Index für `" & zTabName & "`:" & sql & " erstellt", True
     End If
    End If ' not ars.bof
    ars.Close
   Else ' obQMySQL -> no obQMySQL
    sql = vNS
'    IF (td.Attributes AND dbAttachedTable) = 0 AND (td.Attributes AND dbAttachedODBC) = 0 AND (td.Attributes AND dbSystemObject) = 0 THEN
    If td.Type <> "LINK" And td.Type <> "ACCESS TABLE" And td.Type <> "SYSTEM TABLE" Then
    For Each ind In td.Indexes
     obPSZ = ind.PrimaryKey
     obsngibt = False
     If obZMySQL Then
      Set arsz = Nothing
      Set arsz = sAusf("SHOW index FROM `" & zds & "`.`" & zTabName & "` WHERE key_name = '" & ind.name & "'")
      If Not arsz.EOF Then obsngibt = True
      Set arsz = Nothing
     Else
      For i = 0 To zCat.Tables(td.name).Indexes.COUNT - 1
       If zCat.Tables(td.name).Indexes(i).name = ind.name Then
        obsngibt = True
'          IF zcat.Tables(td.Name).Indexes(i).PrimaryKey THEN obPSZ = True
        Exit For
       End If
      Next i
     End If
     obPSZuNE = obPSZ
     If obPSZuNE Then If obai And (aiFName = ind.Columns(0).name) Then obPSZuNE = False
     If obsngibt And Not obPSZuNE Then
       sql = sql & "DROP " & schl(ind.name, obZMySQL)
       If Not obZMySQL Then
        sql = sql & " ON `" & zTabName & "`"
        Call sAusf(sql)
        sql = vNS
       Else
        sql = sql & ","
       End If
     End If
     
     For i = 0 To ind.Columns.COUNT - 1
      Set col = ind.Columns(i)
      If i = 0 Then
       If Not obsngibt Or (obsngibt And Not obPSZuNE) Then
        If obZMySQL Then
         sql = sql & "ADD " & IIf(Not ind.Unique Or obPSZ, vNS, "UNIQUE ") & schl(ind.name, obZMySQL) & "("
        Else
         sql = "CREATE " & IIf(Not ind.Unique, vNS, "UNIQUE ") & "INDEX `" & ind.name & "` ON `" & zTabName & "` ("
        End If
       End If
      End If ' i = 0
      If Not obsngibt Or (obsngibt And Not obPSZuNE) Then
       sql = sql & "`" & col.name & "`"
       zwi = col.name
       If td.Columns(zwi).Type = adLongVarBinary Or td.Columns(zwi).Type = adLongVarChar Or td.Columns(zwi).Type = adLongVarWChar Then
        If obZMySQL Then sql = sql & "(20)"
       End If
       sql = sql & IIf(col.SortOrder = adSortDescending, " DESC", vNS) & ","
      End If
      
      If i = ind.Columns.COUNT - 1 Then
       If LenB(sql) <> 0 Then
        sql = Left$(sql, Len(sql) - 1) & ")"
        If Not obZMySQL Then
'         call sausf ("ALTER TABLE `" & zTabName & "` " & sql)
         Call sAusf(sql & IIf(obPSZ, " WITH PRIMARY", vNS))
         Lese.Ausgeb " Index für `" & zTabName & "`:" & sql & " erstellt", True
         sql = vNS
        Else
         sql = sql & ","
        End If
       End If
      End If
     Next i
    Next ind
    If td.Indexes.COUNT > 0 And obZMySQL Then
     Call sAusf("ALTER TABLE `" & zTabName & "` " & Left$(sql, Len(sql) - 1), , , True)
     Lese.Ausgeb " Index für `" & zTabName & "`:" & sql & " erstellt", True
    End If
    End If
   End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCopyIndices/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCopyIndices

Function doCopyDaten(td As ADOX.Table, cnz As ADODB.Connection, zTabName$, qCat As ADOX.Catalog, obZMySQL%, obQMySQL%)
 Dim rs As New ADODB.Recordset, rsq As New ADODB.Recordset
 Dim PrimFeld$, obdi%
 Dim rsprim As New ADODB.Recordset, fld As Field
 Dim dszahl&, sql2$, lauf%, rAf&, isql$, obprüf%, i&
 Dim xfld As ADOX.Column, f2 As ADOX.Column
 Const mitFeldern% = True
  On Error GoTo fehler
  Set rs = sAusf("SET foreign_key_checks=0", False, rAf) ' vor 27.8.09: obRückg=True
'  Set rs = sAusf("BEGIN", False, rAf) ' vor 27.8.09: obRückg=True
  BegTrans
  If Me.nurSchreiben = 0 Then
   Set rs = sAusf("DELETE FROM `" & DefDB(cnz) & "`.`" & zTabName & "`", False, rAf) ' vor 27.8.09: obRückg=True ' DefDB(cnz)
  Else
   Set rs = sAusf("DELETE FROM `"" & DBn & ""`.`" & zTabName & "`", False, rAf) ' vor 27.8.09: obRückg=True ' DefDB(cnz)
  End If
'  Set rs = sAusf("COMMIT", False, rAf) ' vor 27.8.09: obRückg=True
  ComTrans
  obdi = 0
  If obZMySQL = obQMySQL And Me.nurSchreiben = 0 Then
   If obZMySQL = 0 Or (LCase$(cnz.Properties("Server Name")) = LCase$(DBCn.Properties("Server Name"))) Then
    obdi = True
   End If
  End If
  If obdi Then
'      myEFrag "BEGIN"
      BegTrans
      myEFrag "SET foreign_key_checks=0"
      myEFrag "INSERT INTO `" & DefDB(cnz) & "`.`" & zTabName & "` SELECT * FROM `" & td.name & "`", rAf
      myEFrag "SET foreign_key_checks=1"
'      myEFrag "COMMIT"
      ComTrans
   Else ' obZMySQL = obQMySQL AND Me.nurSchreiben = 0 AND ...
     dszahl = 0
     myFrag rsq, "SELECT * FROM `" & td.name & "`"
     If obQMySQL Then
      myFrag rsprim, "SHOW index FROM `" & td.name & "` WHERE key_name = 'PRIMARY'"
      If Not rsprim.BOF Then
       rsprim.Move 1
       If rsprim.EOF Then
        rsprim.Move -1
        PrimFeld = rsprim!COLUMN_NAME
       End If
      End If
     Else ' obQMySQL
      qCat.ActiveConnection = cnz
      For Each fld In rs.Fields
       If qCat.Tables(td.name).Columns(fld.name).Properties("autoincrement") Then
        PrimFeld = fld.name
        Exit For
       End If
      Next fld
     End If
     Do While Not rsq.EOF
      For lauf = rsq.Fields.COUNT - 1 To 0 Step -1
      sql = "INSERT INTO `" & zTabName & "`" & IIf(mitFeldern, " (", vNS)
      sql2 = " values ("
      For i = 0 To lauf
        Set fld = rsq.Fields(i)
'       IF InStr(fld, "?") <> 0 THEN
'       IF fld.Name = "Größe" OR fld.Name = "Gewicht" THEN
'       IF LCase$(fld.Name) = "augensp zuletzt" AND rsq!Pat_id = 51 THEN
'       IF LCase$(fld.Name) = "schwanger seit" AND rsq!Pat_id = 63 THEN
       If fld.name <> PrimFeld Then
       obprüf = 0
       If IsNull(rsq.Fields(fld.name)) Then
        isql = "NULL"
       ElseIf IsDate(rsq.Fields(fld.name)) And (fld.Type = 7 Or fld.Type = 64 Or fld.Type = 133 Or fld.Type = 134 Or fld.Type = 135) Then
        isql = datformZ(rsq.Fields(fld.name), obZMySQL)
       ElseIf fld.Type = adBoolean And (fld.Type = 11 Or fld.Type = 16 Or fld.Type = 17) Then
        isql = IIf(rsq.Fields(fld.name), "-1", "0")
       ElseIf IsNumeric(rsq.Fields(fld.name)) And (fld.Type <> 8 And fld.Type <> 129 And fld.Type <> 130 And fld.Type <> 200 And fld.Type <> 201 And fld.Type <> 202 And fld.Type <> 203) Then
'        obprüf = True
        isql = REPLACE$(CStr(rsq.Fields(fld.name)), ",", ".")
       Else
        obprüf = True
        If InStrB(rsq.Fields(fld.name), "'") <> 0 Then
         isql = "'" & REPLACE$(rsq.Fields(fld.name), "'", "''") & "'"
        Else
         isql = "'" & rsq.Fields(fld.name) & "'"
        End If
        If InStrB(isql, Chr(0)) <> 0 Then
         isql = REPLACE$(isql, Chr(0), vNS)
        End If
        If obZMySQL Then
         If InStrB(isql, """") <> 0 Then
          isql = REPLACE$(isql, """", "\""")
         End If
         If InStrB(isql, "\") <> 0 Then
          isql = REPLACE$(isql, "\", "\\")
         End If
         If nurSchreiben <> 0 Then
         End If
        Else ' obZMySQL THEN
         If InStrB(isql, vbCrLf) <> 0 Then
          isql = REPLACE$(isql, vbCrLf, vbLf)
         End If
        End If ' obZMySQL THEN
        If nurSchreiben <> 0 Then
         If InStrB(isql, vbCrLf) <> 0 Then
          isql = REPLACE$(isql, vbCrLf, "; ")
         End If
         If InStrB(isql, vbLf) <> 0 Then
          isql = REPLACE$(isql, vbLf, "; ")
         End If
         If InStrB(isql, vbCr) <> 0 Then
          isql = REPLACE$(isql, vbCr, "; ")
         End If
        End If ' obZMySQL
       End If
       If obprüf Then
'        IF Len(isql) - 2 > fld.DefinedSize AND NOT (IsNumeric(rsq.Fields(fld.Name)) AND rsq.Fields(fld.Name) = 255 AND fld.DefinedSize = 1) THEN
'
'        END IF
       End If
       If mitFeldern Then sql = sql & "`" & fld.name & "`, "
       sql2 = sql2 & isql & ", "
       End If ' fld.name <> PrimFeld
      Next i
      dszahl = dszahl + 1
      Lese.Ausgeb "Schreibe Datensatz Nr. " & dszahl & " in Tabelle `" & zTabName & "`", False
'      IF dszahl > 1000 THEN Exit Do
      sql = Left$(sql, Len(sql) - 2) + ")" & Left$(sql2, Len(sql2) - 2) + ")"
      On Error Resume Next
      Err.Clear
      rAf = 0
      Set rs = sAusf(sql, True, rAf)
      If rAf = 1 Then
       Exit For ' lauf
      Else
       If lauf <= 1 Then
        MsgBox "Stop in doCopyDaten durch lauf <= 1 in doCopyDaten, evtl. nicht gelöschte Anamnebogentabelle"
        Stop
       End If
       Exit Function
      End If
      On Error GoTo fehler
      Next lauf
      DoEvents
      rsq.Move 1
     Loop
     Lese.Ausgeb Right$(Space$(10) & dszahl, 10) & " Sätze in Tabelle `" & zTabName & "` geschrieben", True
     rsq.Close
  End If ' obZMySQL = obQMySQL AND Me.nurSchreiben = 0 AND ...
  Call sAusf("SET foreign_key_checks=1", True, rAf)
  Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCopyDaten/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCopyDaten

Private Sub CptListeGanz()
 Dim i
 On Error GoTo fehler
' Me.ServerZ.Width = 3500
'  altCpt = Me.ServerZ
'  Me.ServerZ.Clear
'  changeStill = True
'  Me.ServerZ = altCpt
'  changeStill = False
 Dim ischoda%
 For i = Me.ServerZ.ListCount - 1 To 0 Step -1
  If Me.ServerZ.List(i) = Me.ServerZ.Text And Not ischoda Then
   ischoda = True
  Else
   Me.ServerZ.RemoveItem i
  End If
 Next i
 For Each Comp In Cpts
'  IF Comp <> Me.ServerZ.Text THEN
   Me.ServerZ.AddItem Comp
'  END IF
 Next Comp
 Me.NurLauf.Caption = NL0
 obNurLauf = False
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in CptListeGanz/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' CptListeGanz

Private Sub NurLauf_Click()
 Dim rTs As New ADODB.Connection
' dim rs As New ADODB.Recordset ' geht auch nicht schneller
 Dim i&, otrei$, odb$
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 If obNurLauf Then
  Call CptListeGanz
 Else
'   For i = 0 To Me.ODBC.ListCount
'    odb = Me.ODBC.List(i)
'    IF InStrB(UCase$(odb), "MYSQL") <> 0 THEN
'     otrei = odb
'     Exit For
'    END IF
'   Next i
'  For Each Comp In Cpts
  For i = Me.ServerZ.ListCount - 1 To 0 Step -1
   Set rTs = Nothing
'   SET rs = Nothing
   Err.Clear
   On Error Resume Next
   rTs.Open "DRIVER={" & ODBCStr & "};server=" & Trim$(Left$(Me.ServerZ.List(i), 15)) & ";option=3;uid=praxis;pwd=...;"
'   rs.Open "SELECT * FROM mysql.user LIMIT 1", "DRIVER={" & Me.ODBC & "};server=" & trim$(LEFT(Me.ServerZ.List(i), 15)) & ";option=3;uid=" & Me.uid & ";pwd=" & Me.pwd & ";", adOpenStatic, adLockReadOnly
   If Err.Number = 0 Or InStrB(Err.Description, "denied") > 0 Then ' Access denied
    On Error GoTo fehler
'    Me.ServerZ.AddItem Me.ServerZ.List(i)
   Else
    Me.ServerZ.RemoveItem i
    On Error GoTo fehler
   End If
  Next i
  Me.NurLauf.Caption = NL1
  obNurLauf = True
 End If
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
  ' vermutlich ist kein WMI installiert
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in NurLauf_Click/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' NurLauf_Click

Private Sub ServerZ_DropDown()
' Call Verbind
  If Me.ServerZ.ListCount = 0 Then
   Screen.MousePointer = vbHourglass
   Call ShowAllDomains
   Call ServerZListeGanz
   Screen.MousePointer = vbNormal
  End If
End Sub ' ServerZ_DropDown()

Private Sub ServerZListeGanz()
 Dim i&
 On Error GoTo fehler
' Me.Serverz.Width = 3500
'  altCpt = Me.serverz
'  Me.serverz.Clear
'  changeStill = True
'  Me.serverz = altcpt
'  changeStill = False
 Dim ischoda%
 For i = Me.ServerZ.ListCount - 1 To 0 Step -1
  If Me.ServerZ.List(i) = Me.ServerZ.Text And Not ischoda Then
   ischoda = True
  Else
   Me.ServerZ.RemoveItem i
  End If
 Next i
 For Each Comp In Cpts
'  IF Comp <> Me.ServerZ.Text THEN
   Me.ServerZ.AddItem Comp
'  END IF
 Next Comp
 'Me.NurLauf.Caption = NL0
 obNurLauf = False
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ServerZListeGanz/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ServerZListeGanz

Sub ShowAllDomains()
  Dim oNameSpace  As Object
  Dim oDomain     As Object
 
  On Error GoTo fehler
  Set oNameSpace = GetObject("WinNT:")

  For Each oDomain In oNameSpace
    Debug.Print oDomain.name
    Call ShowAllComputers(oDomain.name)
  Next
  Exit Sub
fehler:
  ' vermutlich ist kein WMI installiert
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ShowAllDomains/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ShowAllDomains

Public Sub ShowAllComputers(ByVal strDomain As String)
  Dim PrimDomainContr     As Object
  Dim oComputer           As Object
  On Error GoTo fehler
  Set PrimDomainContr = GetObject("WinNT://" & strDomain)
  PrimDomainContr.Filter = Array("Computer")
  On Error Resume Next
  
  For Each oComputer In PrimDomainContr
    Cpts.Add Left$(oComputer.name & Space$(CptLänge), CptLänge) & "| " & HostByName(oComputer.name)
    Debug.Print oComputer.name
  Next
  On Error GoTo fehler
  Exit Sub

fehler:
  ' vermutlich ist kein WMI installiert
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ShowAllComputers/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' ShowAllComputers

Function doCopyRelations(obQMySQL%, obZMySQL%, qds$, zds$)
 Dim sql$, sql2$, sql3$
 Dim ars As New ADODB.Recordset, arsz As New ADODB.Recordset
 Dim TabHier$, TabRef$, ConstrNam$
 
 Set arsz = Nothing
 On Error GoTo fehler
 If obQMySQL Then

'  Const constraint_name% = 0, table_name% = 1, referenced_table_name% = 2, referenced_column_name% = 3, column_name% = 4, update_rule% = 5, delete_rule% = 6
  
  Dim qcont(), qi%, qj&
  Screen.MousePointer = vbHourglass
  Lese.Ausgeb "Öffne Tabelle Information-Schema im Quellsystem ...", False
  myFrag arsz, "SELECT r.constraint_name, r.table_name, r.referenced_table_name, referenced_column_name, column_name, update_rule, delete_rule FROM information_schema.referential_constraints r LEFT JOIN information_schema.key_column_usage k USING (constraint_schema, constraint_name, table_name) WHERE constraint_schema = '" & qds & "' ORDER BY constraint_schema, table_name, constraint_name, ordinal_position"
  Screen.MousePointer = vbNormal
  DoEvents
  ReDim qcont(arsz.Fields.COUNT, 0) ' das letzte Feld bedeutet: Schlüssel neu, drop kann entfallen
 
  Do While Not arsz.EOF
   ReDim Preserve qcont(arsz.Fields.COUNT, UBound(qcont, 2) + 1)
   For qi = 0 To arsz.Fields.COUNT - 1
    qcont(qi, UBound(qcont, 2)) = arsz.Fields(qi).Value
   Next qi
   qcont(UBound(qcont, 1), UBound(qcont, 2)) = obZMySQL ' true => kein drop
   arsz.Move 1
  Loop
  arsz.Close
  
  If obZMySQL And cnz.State = 1 And nurSchreiben = 0 Then
     Dim zcont(), zi%, zj&
     Screen.MousePointer = vbHourglass
     Lese.Ausgeb "Öffne Tabelle Information-Schema im Zielsystem...", True
     myFrag arsz, "SELECT r.constraint_name, r.table_name, r.referenced_table_name, referenced_column_name, column_name, update_rule, delete_rule FROM information_schema.referential_constraints r LEFT JOIN information_schema.key_column_usage k USING (constraint_schema, constraint_name, table_name) WHERE constraint_schema = '" & zds & "' ORDER BY constraint_schema, table_name, constraint_name", adOpenStatic, cnz, adLockReadOnly
     Screen.MousePointer = vbNormal
     DoEvents
     ReDim zcont(arsz.Fields.COUNT - 1, 0)
    
     Do While Not arsz.EOF
      ReDim Preserve zcont(arsz.Fields.COUNT - 1, UBound(zcont, 2) + 1)
      For zi = 0 To arsz.Fields.COUNT - 1
       zcont(zi, UBound(zcont, 2)) = arsz.Fields(zi).Value
      Next zi
      arsz.Move 1
     Loop
     arsz.Close
  End If

'  SET ars = Nothing
'  myFrag ars, "SELECT * FROM information_schema.referential_constraints WHERE constraint_schema = '" & qds & "'"
'  Do While Not ars.EOF
  For qj = 1 To UBound(qcont, 2)
' Prüfen, ob nicht schon auf dem Zielsystem genauso vorhanden
   Dim schonda%, ZZ%
   schonda = 0
'   TabHier = ars!table_name
'   TabRef = ars!referenced_table_name
   TabHier = qcont(table_name, qj)
   TabRef = qcont(referenced_table_name, qj)
   ConstrNam = qcont(constraint_name, qj)
   If obZMySQL And cnz.State = 1 And nurSchreiben = 0 Then
    zj = 1
    Do
     If zj > UBound(zcont, 2) Then Exit Do
     If zcont(constraint_name, zj) = qcont(constraint_name, qj) Then qcont(UBound(qcont, 1), qj) = 0 ' dann doch drop
     If zcont(table_name, zj) = qcont(table_name, qj) And zcont(referenced_table_name, zj) = qcont(referenced_table_name, qj) And zcont(constraint_name, zj) = qcont(constraint_name, qj) Then
      schonda = True
      ZZ = 0
      Do
       If zcont(table_name, zj + ZZ) <> qcont(table_name, qj + ZZ) Or zcont(referenced_table_name, zj + ZZ) <> qcont(referenced_table_name, qj + ZZ) Or zcont(constraint_name, zj + ZZ) <> qcont(constraint_name, qj + ZZ) Or zcont(referenced_column_name, zj + ZZ) <> qcont(referenced_column_name, qj + ZZ) Or zcont(COLUMN_NAME, zj + ZZ) <> qcont(COLUMN_NAME, qj + ZZ) Or zcont(update_rule, zj + ZZ) <> qcont(update_rule, qj + ZZ) Or zcont(delete_rule, zj + ZZ) <> qcont(delete_rule, qj + ZZ) Then
        schonda = 0
        Exit Do
       End If
       ZZ = ZZ + 1
       If qj + ZZ > UBound(qcont, 2) Or zj + ZZ > UBound(zcont, 2) Then Exit Do
       If qcont(table_name, qj + ZZ) <> TabHier Or qcont(referenced_table_name, qj + ZZ) <> TabRef Or qcont(constraint_name, qj + ZZ) <> ConstrNam Then Exit Do
      Loop
      If schonda Then qj = qj + ZZ - 1
      Exit Do
     End If
     zj = zj + 1
    Loop
   End If
   
   If Not schonda Then
   
       sql2 = vNS
    '   sql = sql & "ADD CONSTRAINT `" & ars!constraint_name & "` FOREIGN KEY ("
       sql = sql & "ADD CONSTRAINT `" & qcont(constraint_name, qj) & "` FOREIGN KEY ("
       
       
    '   SET arsz = Nothing
    '   myFrag arsz, "SELECT * FROM information_schema.key_column_usage WHERE constraint_schema = '" & qds & "' AND constraint_name = '" & ars!constraint_name & "' AND table_name = '" & TabHier & "'"
        sql3 = vNS
    '   Do While Not arsz.EOF
       Do
    '    sql2 = sql2 & "`" & arsz!referenced_column_name & "`,"
    '    sql3 = sql3 & "`" & arsz!column_name & "`,"
    '    arsz.Move 1
        sql2 = sql2 & "`" & qcont(referenced_column_name, qj) & "`,"
        sql3 = sql3 & "`" & qcont(COLUMN_NAME, qj) & "`,"
        qj = qj + 1
        If qj > UBound(qcont, 2) Then
         qj = qj - 1
         Exit Do
        End If
        If qcont(referenced_table_name, qj) <> TabRef Or qcont(table_name, qj) <> TabHier Or qcont(constraint_name, qj) <> ConstrNam Then
         qj = qj - 1
         Exit Do
        End If
       Loop
       
       If obZMySQL Then
        TabHier = LCase$(TabHier)
        TabRef = LCase$(TabRef)
        ConstrNam = LCase$(ConstrNam)
       End If
       
       sql3 = Left$(sql3, Len(sql3) - 1)
    '   sql = sql & sql3 & ") " & "references `" & TabRef & "` (" & LEFT(sql2, Len(sql2) - 1) & ")" & IIf(obZMySQL, " ON update " & ars!update_rule & " ON delete " & ars!delete_rule, vns)
       sql = sql & sql3 & ") " & "REFERENCES `" & TabRef & "` (" & Left$(sql2, Len(sql2) - 1) & ")" & IIf(obZMySQL, " ON update " & qcont(update_rule, qj) & " ON DELETE " & qcont(delete_rule, qj), vNS)
    '   ON Error Resume Next
    '   call sausf ("ALTER TABLE `" & TabHier & "` " & "drop " & IIf(obZMySQL, "FOREIGN KEY", "Constraint") & "`" & ars!constraint_name & "`")
       If qcont(UBound(qcont, 1), qj) = 0 Then Call sAusf("ALTER TABLE `" & TabHier & "` " & "drop " & IIf(obZMySQL, "FOREIGN KEY", "Constraint") & "`" & qcont(constraint_name, qj) & "`", , , True)
       Err.Clear
       Call sAusf("ALTER TABLE `" & TabHier & "` " & sql, , , True)
       Debug.Print IIf(ErrNumber = 0, "Kein Fehler", "Err.Nr " & ErrNumber & ", " & ErrDescr) & " bei " & sql
       If ErrNumber = -2147467259 And Err.Description = "Es wurde kein eindeutiger Index für das in Beziehung stehende Feld der Primärtabelle angegeben." Then
       ' hier dürfte er nicht mit nurSchreiben hinkommen
        Err.Clear
        Call sAusf("create unique index `" & TabHier & "_rel` ON `" & TabRef & "` (" & Left$(sql2, Len(sql2) - 1) & ")")
        Debug.Print Err.Number, Err.Description
        On Error GoTo fehler
        Call sAusf("ALTER TABLE `" & TabHier & "` " & sql, , , True)
       End If
       On Error GoTo fehler
       sql = vNS
   End If 'not schonda
  Next qj
'   ars.Move 1
'  Loop
'  ars.Close
  If arsz.State = 1 Then arsz.Close
 Else ' obQMySQL => nicht
  Set ars = Nothing
  On Error Resume Next
  Dim ErrNr&, ErrDes$
  Do
   Err.Clear
   myFrag ars, "SELECT * FROM msysrelationships", , , , , , True, ErrNr, ErrDes
   If ErrNr <> 0 Then
    MsgBox ErrDes & vbCrLf & "Falls Berechtigungsfehler, dann bitte in Access Berechtigungen setzten: Extras -> Sicherheit -> Benutzer- und Gruppenberechtigungen (admin für mysysrelationships)"
   Else
    Exit Do
   End If
  Loop
  On Error GoTo fehler
  Do While Not ars.EOF
   sql2 = vNS
   sql = sql & "ADD CONSTRAINT `" & ars!szrelationship & "` FOREIGN KEY ("
   
   TabHier = ars!szobject
   If obZMySQL Then TabHier = LCase$(TabHier)
   TabRef = ars!szreferencedobject
   If obZMySQL Then TabRef = LCase$(TabRef)
   
   Set arsz = Nothing
   myFrag arsz, "SELECT * FROM msysrelationships WHERE szrelationship = '" & ars!szrelationship & "' AND szobject = '" & TabHier & "'"
   sql3 = vNS
   Do While Not arsz.EOF
    sql2 = sql2 & "`" & arsz!szreferencedcolumn & "`,"
    sql3 = sql3 & "`" & arsz!szcolumn & "`,"
    arsz.Move 1
   Loop
   
   sql3 = Left$(sql3, Len(sql3) - 1)
   sql = sql & sql3 & ") " & "REFERENCES `" & TabRef & "` (" & Left$(sql2, Len(sql2) - 1) & ")" & IIf(obZMySQL, " ON UPDATE CASCADE ON DELETE RESTRICT", vNS)
   On Error Resume Next
   Call sAusf("ALTER TABLE `" & TabHier & "` " & "drop " & IIf(obZMySQL, "FOREIGN KEY", "CONSTRAINT") & "`" & ars!szrelationship & "`")
   Err.Clear
   Call sAusf("ALTER TABLE `" & TabHier & "` " & sql)
   If Err.Number = -2147467259 And Err.Description = "Es wurde kein eindeutiger Index für das in Beziehung stehende Feld der Primärtabelle angegeben." Then
    Err.Clear
    Call sAusf("CREATE UNIQUE INDEX `" & TabHier & "_rel` ON `" & TabRef & "` (" & Left$(sql2, Len(sql2) - 1) & ")")
    Debug.Print Err.Number, Err.Description
    On Error GoTo fehler
    Call sAusf("ALTER TABLE `" & TabHier & "` " & sql)
   ElseIf Err.Number <> 0 Then
    Debug.Print Err.Number, Err.Description
   End If
   On Error GoTo fehler
   sql = vNS
   ars.Move 1
  Loop
  ars.Close
  arsz.Close
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doCopyRelations/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' doCopyRelations

Function sAusf(sql$, Optional obRückg%, Optional rAf&, Optional obtolerant%, Optional Einzug%) As ADODB.Recordset
 Dim FMeld$, FNr&, nsql As New CString, i&, nskurz$
 Static obüf%
 ErrNumber = 0
 On Error GoTo fehler
 If Me.nurSchreiben = 0 Then sql = REPLACE$(sql, """", """""") ' 16.8.09: Weiß nicht, obs das braucht
 If InStrB(sql, "¤¤") <> 0 Then
  If Me.nurSchreiben <> 0 Then
   sql = REPLACE$(sql, "¤¤", """") ' ASCII 207
  Else '
   sql = REPLACE$(sql, "¤¤ & DBN & ¤¤", Me.DBn) ' ASCII 207
  End If
 End If
 If Me.nurSchreiben <> 0 Then
  On Error Resume Next
  Open Me.SchreibenAuf For Append As #299
  If Err.Number <> 0 Then
   MsgBox " Fehler " & Err.Number & ",: " & Err.Description & vbCrLf & "beim Schreibersucht auf " & Me.SchreibenAuf & vbCrLf & "breche ab."
   rAf = -1
   Exit Function
  End If
  On Error GoTo fehler
  nsql = sql
  If nsql.length <= 5000 Then
   If nsql.length <= 900 Then
    If Right$(nsql, 5) = " & "" " Then
     nskurz = nsql.Left(nsql.length - 5) & ","
    Else
     nskurz = nsql & ""","
    End If
    Print #299, Space$(Einzug) & " call doex(""" & nskurz & IIf(obtolerant, "-1", "0") & ")"
   Else
    Print #299, Space$(Einzug) & " call doex(""" & nsql.Left(900) & """ & _ "
    Set nsql = nsql.Mid(901)
    Do
     If nsql.length <= 900 Then
      Print #299, Space$(Einzug) & """" & nsql & """," & IIf(obtolerant, "-1", "0") & ")"
      Exit Do
     Else
      Print #299, Space$(Einzug) & """" & nsql.Left(900) & """ & _ "
      Set nsql = nsql.Mid(901)
     End If
    Loop
   End If
  Else
   If Not obüf Then
    Print #299, Space$(Einzug) & " dim nsql$"
    obüf = True
   End If
   Print #299, Space$(Einzug) & " nsql="""""
   Do
    If Len(sql) <= 900 Then
     Print #299, Space$(Einzug) & " nsql = nsql & """ & nsql.Left(900) & """"
     Set nsql = nsql.Mid(901)
     Exit Do
    Else
     Print #299, Space$(Einzug) & " nsql = nsql & """ & nsql.Left(900) & """ & _ "
     Set nsql = nsql.Mid(901)
     For i = 0 To 5
      If Len(nsql) <= 900 Or i = 5 Then
       Print #299, Space$(Einzug) & """" & nsql.Left(900) & """"
       Set nsql = nsql.Mid(901)
       Exit For
      Else
       Print #299, Space$(Einzug) & """" & nsql.Left(900) & """ & _ "
       Set nsql = nsql.Mid(901)
      End If
     Next i
     If nsql.length = 0 Then Exit Do
    End If
   Loop
   Print #299, Space$(Einzug) & " call doex(nsql," & IIf(obtolerant, "-1", "0") & ")"
  End If ' Len(nsql) <= 5000 THEN
  Close #299
 End If ' Me.nurSchreiben <> 0 THEN
 If Me.nurSchreiben = 0 Or obRückg Then
  On Error Resume Next
  myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", , cnz
  Err.Clear
  If cnz.State = 0 Then
   Call setzmpwd
   cnz.Open cnzCStr & mpwd & ";"
   ErrNumber = Err.Number
   ErrDescr = Err.Description
   MsgBox "Fehler " & ErrNumber & " beim Öffnen von '" & cnz & "'" & vbCrLf & "Beschreibung: " & Err.Description
   ProgEnde
  End If
  Set sAusf = myEFrag(sql, rAf, cnz, True, ErrNumber, ErrDescr)
'  ErrNumber = Err.Number
  FMeld = IIf(ErrNumber = 0, "Kein Fehler", "Err.Nr " & ErrNumber) & ", rAf: " & rAf & " bei " & sql & "; " & vbCrLf & "   " & ErrDescr
  Debug.Print "sAusf, FMeld: " & FMeld
  If FNr <> 0 And (InStrB(LCase$(sql), "drop") = 0 Or InStrB(LCase$(sql), ",") <> 0) Then
   Lese.Ausgeb FMeld, True
   MsgBox "Stop in sAusf durch Fnr<>0 AND (InStrB(LCase$(sql), ""drop"") = 0 OR InStrB(LCase$(sql), "","") <> 0)" & vbCrLf & "sql: " & sql
   Stop
  End If
  If Me.nurSchreiben = 0 Then Print #302, FMeld
 Else
  rAf = 1
 End If
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescr = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 'Call XPH
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), vNS, CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescr + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in sAusf/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' sAusf(sql$, obRückg%) As Adodb.Recordset

' schreibt Befehle auf für spätere Ausführung, im Gegensatz zu sAusf, das alternativ für sofortige Ausführung zur Verfügugn steht
' => alle Aufrufe für Ausf müßten die Alternative sofortige Ausführung anderweitig zur Verfügung stellen
Function Ausf(ByRef Befehl$)
 Const Einzug& = 2
 Dim nbefehl As New CString, nskurz$
 If Me.nurSchreiben <> 0 Then
  Open Me.SchreibenAuf For Append As #299
'  Print #299, befehl
  nbefehl = Befehl
'  IF nbefehl.Length <= 5000 THEN
   If nbefehl.length <= 900 Then
    Print #299, nbefehl
   Else
    Print #299, nbefehl.Left(900) & """ & _ "
    Set nbefehl = nbefehl.Mid(901)
    Do
     If nbefehl.length <= 900 Then
      Print #299, Space$(Einzug) & """" & nbefehl
      Exit Do
     Else
      Print #299, Space$(Einzug) & """" & nbefehl.Left(900) & """ & _ "
      Set nbefehl = nbefehl.Mid(901)
     End If
    Loop
   End If
'  Else
'   IF Not obüf THEN
'    Print #299, Space$(Einzug) & " dim nbefehl$"
'    obüf = True
'   END IF
'   Print #299, Space$(Einzug) & " nbefehl="""""
'   Do
'    IF Len(sql) <= 900 THEN
'     Print #299, Space$(Einzug) & " nbefehl = nbefehl & """ & nbefehl.Left(900) & """"
'     SET nbefehl = nbefehl.Mid(901)
'     Exit Do
'    Else
'     Print #299, Space$(Einzug) & " nbefehl = nbefehl & """ & nbefehl.Left(900) & """ & _ "
'     SET nbefehl = nbefehl.Mid(901)
'     For i = 0 To 5
'      IF Len(nbefehl) <= 900 OR i = 5 THEN
'       Print #299, Space$(Einzug) & """" & nbefehl.Left(900) & """"
'       SET nbefehl = nbefehl.Mid(901)
'       Exit For
'      Else
'       Print #299, Space$(Einzug) & """" & nbefehl.Left(900) & """ & _ "
'       SET nbefehl = nbefehl.Mid(901)
'      END IF
'     Next i
'     IF nbefehl.Length = 0 THEN Exit Do
'    END IF
'   Loop
'   Print #299, Space$(Einzug) & " call doex(nbefehl," & IIf(obtolerant, "-1", "0") & ")"
'  END IF ' Len(nbefehl) <= 5000 THEN
  Close #299
 End If
End Function ' Ausf(befehl)

' in dbKopier, dbCopyAllMyMy
Function SchreibF1(obQMySQL%)
  Dim Server$, spos&, sp2&
  Open Me.SchreibenAuf For Output As #299
  If obQMySQL Then Server = GetServr(DBCn)
  Print #299, "'Bauanleitung für eine Datenbank wie `" & IIf(obQMySQL, "//" & Server & "/", vNS) & DefDB(DBCn) & "` vom " & Format(Now(), "d.m.yy hh:mm:ss")
  Print #299, "Option Explicit"
'  Print #299, "Const DBn$ = """ & DBn & """ ' Datenbankname"
  Print #299, "Dim cnzCStr$ ' da unter Vista der Connectionstring jetzt nicht mehr aussagekräftig ist"
  Print #299, "Dim cnz As New ADODB.connection, FNr&, lErrNr& ' letzter Fehler bei doEx"
  Print #299, "Dim obProt% ' ob Protokollierung stattfindet, da Protokolldatei zu öffnen"
  Close #299
End Function ' SchreibF1

' in dbKopier, dbCopyAllMyMy
Function SchreibF2(DBn$)
  Open Me.SchreibenAuf For Append As #299
  Print #299, ""
  Print #299, "Function doEx&(sql$, obtolerant%) ' SQL-Befehl ausführen, Fehler anzeigen"
  Print #299, " Dim rAF&, FMeld$"
  Print #299, " Dim lErrNr&, fDesc$"
  Print #299, " On Error Resume Next"
  Print #299, " cnz.DefaultDatabase = hDBn"
  Print #299, " IF obtolerant THEN ON Error Resume Next ELSE ON Error GoTo fehler"
  Print #299, " myEFrag sql, rAf, cnz, True, lErrNr, fDesc"
  Print #299, "' lErrNr = Err.Number"
  Print #299, " FMeld = IIf(lErrNr = 0, ""Kein Fehler"", ""Err.Nr "" & lErrNr & "", "" & fDesc) & "", rAf: "" & rAF & "" bei "" & sql"
  Print #299, " ON Error GoTo fehler"
  Print #299, " Debug.Print ""doEx, FMeld: "" & FMeld"
  Print #299, " Debug.Print fDesc"
  Print #299, " IF obProt THEN Print #302, FMeld"
  Print #299, " DoEvents"
  Print #299, " Exit Function"
  Print #299, "fehler:"
  Print #299, " Dim AnwPfad$"
  Print #299, "#If VBA6 THEN"
  Print #299, " AnwPfad = currentDB.Name"
  Print #299, "#Else"
  Print #299, " AnwPfad = App.Path"
  Print #299, "#END IF"
  Print #299, "SELECT CASE Err.Number"
  Print #299, " Case -2147467259 "
  Print #299, "  IF InStrB(Err.Description, ""nicht erzeugen"") THEN ' 'Kann Tabelle 'testDB1.faxe' nicht erzeugen (Fehler: 150)"
  Print #299, "   doEx = 150"
  Print #299, "   Exit Function"
  Print #299, "  ElseIf InStrB(Err.Description, ""is not BASE TABLE"") <> 0 THEN"
  Print #299, "   doEx = 151"
  Print #299, "   Exit Function"
  Print #299, "  ElseIf InStrB(Err.Description, ""MySQL server has gone away"") THEN"
  Print #299, "   cnz.Close"
  Print #299, "   cnz.Open"
  Print #299, "   Call doEx(""USE `"" & hDBn & ""`"", 0)"
  Print #299, "   Resume"
  Print #299, "  END IF"
  Print #299, "End SELECT"
  Print #299, "SELECT CASE MsgBox(""FNr: "" & FNr & "", ErrNr: "" & CStr(Err.Number) & vbCrLf & ""LastDLLError: "" & CStr(Err.LastDllError) & vbCrLf & ""Source: "" & IIf(ISNULL(Err.source), """", CStr(Err.source)) & vbCrLf & ""Description: "" & Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in doEx/"" & AnwPfad)"
  Print #299, " Case vbAbort: Call MsgBox("" Höre auf ""): Progende"
  Print #299, " Case vbRetry: Call MsgBox("" Versuche nochmal ""): Resume"
  Print #299, " Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
  Print #299, "End SELECT"
  Print #299, "End FUNCTION ' doEx"
  Print #299, ""
  Print #299, "Function SplitN&(ByRef q$, Sep$, erg$()) ' da Split() Speicher fraß"
  Print #299, " Dim p1&, p2&, Slen&, obExit%, runde&"
  Print #299, " ON Error GoTo fehler"
  Print #299, " IF NOT ISNULL(q) THEN"
  Print #299, "  Slen = Len(Sep)"
  Print #299, "  For runde = 1 To 2"
  Print #299, "   p2 = 0"
  Print #299, "   Do"
  Print #299, "    p1 = p2"
  Print #299, "    p2 = InStr(p1 + Slen, q, Sep)"
  Print #299, "    IF p2 = 0 THEN p2 = Len(q) + 1: obExit = True"
  Print #299, "    IF p2 <> 0 THEN"
  Print #299, "     IF runde = 2 THEN"
  Print #299, "      erg(SplitN) = Mid$(q, p1 + Slen, p2 - p1 - Slen)"
  Print #299, "     END IF"
  Print #299, "     SplitN = SplitN + 1"
  Print #299, "    END IF"
  Print #299, "    IF obExit THEN Exit Do"
  Print #299, "   Loop"
  Print #299, "   IF runde = 1 THEN"
  Print #299, "    ReDim erg(SplitN - 1)"
  Print #299, "    SplitN = 0"
  Print #299, "    obExit = 0"
  Print #299, "   END IF"
  Print #299, "  Next runde"
  Print #299, " END IF"
  Print #299, " Exit Function"
  Print #299, "fehler:"
  Print #299, " Dim AnwPfad$"
  Print #299, "#If VBA6 THEN"
  Print #299, " AnwPfad = currentDB.Name"
  Print #299, "#Else"
  Print #299, " AnwPfad = App.Path"
  Print #299, "#END IF"
  Print #299, "SELECT CASE MsgBox(""FNr: "" & FNr & "", ErrNr: "" & CStr(Err.Number) + vbCrLf + ""LastDLLError: "" + CStr(Err.LastDllError) + vbCrLf + ""Source: "" + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + ""Description: "" + Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in SplitN/"" + AnwPfad)"
  Print #299, " Case vbAbort: Call MsgBox(""Höre auf""): Progende"
  Print #299, " Case vbRetry: Call MsgBox(""Versuche nochmal""): Resume"
  Print #299, " Case vbIgnore: Call MsgBox(""Setze fort""): Resume Next"
  Print #299, "End SELECT"
  Print #299, "End FUNCTION ' aufSplit"
  Print #299, ""
  Print #299, "' in calldoGenMachDB_Click"
  Print #299, "Public FUNCTION doMach_" & REPLACE$(DBn, " ", "_") & "(DBn$, DBCn AS ADODB.Connection, Optional Server$, Optional obStumm%=True) ' Datenbankname"
  Print #299, " Dim rsc As New ADODB.Recordset, sct$, Spli$(), tStr$, TMt As New CString, TabEig$, fsql$"
  Print #299, " Dim i&, p1&, p2&, p3&, CLen&, CLen1&, obLT%, ep$, pwp&, pwd$, mpwd$"
  Print #299, " Dim Index$()"
  Print #299, " ON Error Resume Next"
  Print #299, " hDBn = DBn"
  Print #299, " Open App.Path & ""\MachDB.bas_prot.txt"" For Output AS #302"
  Print #299, " obProt = (Err.Number = 0)"
  Print #299, " ON Error GoTo fehler"
  Print #299, " ep = DBCn.Properties(""Extended Properties"")"
  Print #299, " pwp = InStr(UCase$(ep), ""PWD="") + 4"
  Print #299, " If pwp <> 0 Then pwd = Mid$(ep, pwp, InStr(pwp, ep, "";"") - pwp)"
  Print #299, " mpwd = MachDatenbank.setzmpwd()"
  Print #299, " If mpwd = """" Then exit function"
  Close #299
End Function ' SchreibF2

' in Start_Click(MachDatenbank), Machalle_click(MachDatenbank), doMachDB(MachDatenbank)
Function dbKopier(cnz As ADODB.Connection, cnzCStr$, DBn$, Optional obmitDaten%, Optional anschließendverknüpfen%, Optional obohnefertig%) ' Transferiert die nicht verknüpften Tabellen aus der aktuellen Accessdatenbank in eine neu zu erstellende MySQL-Datenbank
 Dim sql$, mft$, i&, obQMySQL%, obZMySQL%, obTabGanzKop%, erg%
' Dim CommentGleich%
 Dim td As ADOX.Table
 Dim Cn As New ADODB.Connection
 Dim zCat As New ADOX.Catalog ' SET zcat = CreateObject("ADOX.Catalog")
 Dim qCat As New ADOX.Catalog ' Ursprünglicher Katalog
 Dim zds$, qds$
 Dim aiFName$, obai%, zTabName$
 Dim runde%
 
 On Error GoTo fehler
 FNr = 1
 If Me.nurSchreiben = 0 Then Open Me.SchreibenAuf & "_prot.txt" For Output As #302
 FNr = 0
 If Me.alleDaten <> 0 Or Me.nurCustomizing <> 0 Then
  If MsgBox("Wollen Sie wirklich alle " & IIf(Me.nurCustomizing <> 0, "Customizing-", vNS) & "Daten kopieren?", vbYesNo + vbDefaultButton2, "Sicherheits-Rückfrage") = vbNo Then
   Unload Me
   Exit Function
  End If
 End If
 Close #256
 FNr = 2
 Open uVerz & "dbKopier.txt" For Output As #256
 FNr = 0
 obZMySQL = IIf(InStrB(cnzCStr, "MySQL") <> 0, True, False)
 obQMySQL = LVobMySQL
 
 qCat.ActiveConnection = DBCn
 qds = DefDB(DBCn)
 
 myEFrag ("SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ")

 If Me.MitTabellen <> 0 Or Me.nurCustomizing <> 0 Or Me.alleDaten <> 0 Then
  If obQMySQL And obZMySQL Then
   Call doCopyAllMyMy(cnz, qCat, zCat)
  Else
   If Me.nurSchreiben <> 0 Then
    Call SchreibF1(obQMySQL)
    Call SchreibF2(DefDB(DBCn))
   End If ' Me.nurSchreiben <> 0 Then
   
   If Me.MitTabellen <> 0 Then Call doMachZielDatenbank(cnz, DBCn, DBn, zCat, obZMySQL)
   If Not zCat.ActiveConnection Is Nothing Then
    zds = DefDB(cnz)
   Else
    zds = DBn
   End If
 
   On Error Resume Next
   If obZMySQL Then
'    Call sAusf("BEGIN WORK")
    Call sAusf("SET FOREIGN_KEY_CHECKS = 0")
   Else
    If Me.nurSchreiben = 0 Then
'     Call cnz.BeginTrans
      BegTrans
    Else
'     Call Ausf(" cnz.begintrans")
     Call Ausf(" BegTrans")
    End If
'    If cnz.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 1
   End If
 
   On Error GoTo fehler
 
   If Me.nurSchreiben <> 0 Then
    Ausf " SET zCat = Nothing ' runde: " & runde
    Ausf " SET zCat.ActiveConnection = cnz"
   End If
   For runde = 1 To 4
    For Each td In qCat.Tables
'    cnz.BeginTrans
      If obZMySQL Then zTabName = LCase$(td.name) Else zTabName = td.name
      If Me.MitTabellen <> 0 Then
       erg = 0
       Select Case td.Type
        Case "VIEW"
         If (runde = 2 And Left$(zTabName, 2) = "__") Or (runde = 3 And Left$(zTabName, 2) <> "__" And Left$(zTabName, 1) = "_") Or (runde = 4 And Left$(zTabName, 1) <> "_") Then
          erg = doCopyView(qds, td.name, cnz, zTabName, obQMySQL, qCat)
         End If
        Case "TABLE"
         If runde < 3 Then
          Dim obLink%
          If Not obQMySQL Then
           obLink = td.Properties("Jet OLEDB:Create Link")
          End If ' Not obQMySQL THEN
          If Not obLink Then
           erg = doCopyTable(td, cnz, zTabName, obQMySQL, obZMySQL, qCat, zCat, runde, aiFName, obai, obTabGanzKop)
           If Not obTabGanzKop Then
            If Me.mitIndices Then ' Indices
             If runde = 2 And Not (obQMySQL And obZMySQL And Not Me.nurSchreiben) Then
'             IF zTabName = "augenbefunde" THEN Stop
              Call doCopyIndices(td, cnz, zTabName, obQMySQL, obZMySQL, zCat, aiFName, obai, qds, zds)
             End If
            End If ' me.mitIndices
           End If ' not obTabGanzKop
          End If ' not obLink
         End If ' runde < 3
        Case "LINK", "ACCESS TABLE", "SYSTEM TABLE"
        Case Else
         MsgBox "Stop in doKopier durch: td.type = " & td.Type
         Stop
       End Select
  '   cnz.CommitTrans
       If erg = 1 Then Lese.Ausgeb Left$(zTabName & Space$(25), 25) & " erstellt in " & zds & " aus " & qds & " Runde: " & runde, True
      End If ' me.mitTabellen
    Next td
   Next runde
  End If ' obQMySQL AND obZMySQL THEN
   
'  END IF ' zTabName = medplan
  If Me.alleDaten <> 0 Or Me.nurCustomizing <> 0 Then
   For Each td In qCat.Tables
    Dim doCop%
'    IF runde = 4 THEN
     doCop = False
     Select Case td.Type
      Case "TABLE"
       If Me.alleDaten <> 0 Then doCop = True
       If Me.nurCustomizing <> 0 Then
        Select Case LCase$(td.name)
         Case "diagg1", "diagreihe", "br_abgehakt", "ebm2000plus", "eintragszahlen", _
             "hausaerzte", "kassenliste", "laborgruppen", "liuez", "medarten", _
             "pauschalen", "werte_scheingruppen", "werte_weggeldzonen"
          doCop = True
         Case "laborxbakt", "laborxeingel", "laborxleist", "laborxsaetze", "laborxus", "laborxwert", _
          "laborybakt", "laborydat", "laboryleist", "laborysaetze", "laboryus", "laborywert"
          If Me.auchLaborX <> 0 Then doCop = True
         Case "anamnesebogen"
          If Me.auchAnamnese <> 0 Then doCop = True
        End Select
       End If
     End Select
     If LCase$(td.name) = "genehmigungen" Then doCop = True
     If doCop Then
      If obZMySQL Then zTabName = LCase$(td.name) Else zTabName = td.name
      Call doCopyDaten(td, cnz, zTabName, qCat, obZMySQL, obQMySQL)
     End If
'    END IF ' runde = 4
   Next td
  End If
 End If ' Me.MitTabellen OR Me.nurCustomizing OR Me.alleDaten THEN
 
' #END IF
' Verknüpfte Tabellen erstellen und formatieren
' IF anschließendverknüpfen THEN Call machVerknTab(DBn)
 
 If Not (obQMySQL And obZMySQL) Then
  If Me.mitRelationen Then
   Call doCopyRelations(obQMySQL, obZMySQL, qds, zds)
  End If ' me.mitRelationen
 End If
' For Each td In qCat.Tables
'  SELECT CASE td.Type
'   Case "TABLE"
'   Dim rel AS ADOX.key
'   For Each rel In td.Keys
'    Debug.Print rel.Name, rel.RelatedTable, rel.Type, rel.UpdateRule
'    For Each col In rel.Columns
'     Debug.Print "         ", col.Name, col.Type, col.RelatedColumn, col.DefinedSize, col.NumericScale, col.Attributes, col.SortOrder, col.Properties.Count
'    Next col
'   Next rel
'   END SELECT
' Next td
 
 If obZMySQL Then
  Call sAusf("SET FOREIGN_KEY_CHECKS = 1")
 End If
 If cnz.State = 1 Then
  On Error Resume Next
  If cnz.DefaultDatabase <> DBCn.DefaultDatabase Or ((cnz.DefaultDatabase = DBCn.DefaultDatabase) And obTrans <> 0) Then ComTrans cnz ' cnz.CommitTrans: If cnz.DefaultDatabase = DBCn.DefaultDatabase Then obTrans = 0
  On Error GoTo fehler
 End If
 If Me.nurSchreiben <> 0 Then
  Open Me.SchreibenAuf For Append As #299
  If Not obZMySQL Then
   Print #299, " IF cnz.DefaultDatabase <> DBCN.DefaultDatabase OR ((cnz.DefaultDatabase = DBCN.DefaultDatabase) AND obTrans <> 0) THEN ComTrans cnz ' cnz.commitTrans: IF cnz.DefaultDatabase = DBCN.DefaultDatabase THEN obTrans = 0"
  End If
  Print #299, " IF obProt THEN Close #302"
  Print #299, " IF not obstumm THEN"
  Print #299, "  MsgBox ""Fertig mit doMach_" & REPLACE$(DefDB(DBCn), " ", "_") & "("" & DBn & "", DBCn,"" & Server & "")!"
  Print #299, " END IF"
  Print #299, " Exit Function"
  Print #299, "fehler:"
  Print #299, " Dim AnwPfad$"
  Print #299, "#If VBA6 THEN"
  Print #299, " AnwPfad = currentDB.Name"
  Print #299, "#Else"
  Print #299, " AnwPfad = App.Path"
  Print #299, "#END IF"
  Print #299, "SELECT CASE MsgBox(""FNr: "" & FNr & "", ErrNr: "" & CStr(Err.Number) & vbCrLf & ""LastDLLError: "" & CStr(Err.LastDllError) & vbCrLf & ""Source: "" & IIf(ISNULL(Err.source), """", CStr(Err.source)) & vbCrLf & ""Description: "" & Err.Description, vbAbortRetryIgnore, ""Aufgefangener Fehler in doMach_" & REPLACE$(DefDB(DBCn), " ", "_") & "/"" & AnwPfad)"
  Print #299, " Case vbAbort: Call MsgBox("" Höre auf ""): Progende"
  Print #299, " Case vbRetry: Call MsgBox("" Versuche nochmal ""): Resume"
  Print #299, " Case vbIgnore: Call MsgBox("" Setze fort ""): Resume Next"
  Print #299, "End SELECT"
  Print #299, "End FUNCTION 'doMach_" & REPLACE$(DefDB(DBCn), " ", "_") & ""
  Print #299, ""
  Print #299, "Function GetServr$(DBCn AS ADODB.Connection)"
  Print #299, " Dim spos&, sp2&, DBCs$"
  Print #299, " DBCs = DBCn.Properties(""Extended Properties"")"
  Print #299, " spos = InStr(1,DBCs, ""server="",vbTextCompare)"
  Print #299, " IF spos <> 0 THEN"
  Print #299, "  sp2 = InStr(spos, DBCs, "";"")"
  Print #299, "  IF sp2 = 0 THEN sp2 = Len(DBCs)"
  Print #299, "  GetServr = Mid$(DBCs, spos + 7, sp2 - spos - 7)"
  Print #299, " END IF ' spos <> 0 THEN"
  Print #299, "End FUNCTION ' GetServr"
  Print #299, ""
  Print #299, "Function AIoZ(Ursp) AS CString ' Ursp kann $ oder CString sein"
  Print #299, " Const Such$ = ""AUTO_INCREMENT="""
  Print #299, " SET AIoZ = New CString"
  Print #299, " AIoZ = Ursp"
  Print #299, " Dim p0&, p1&"
  Print #299, " p0 = AIoZ.Instr(Such)"
  Print #299, " IF p0 <> 0 THEN"
  Print #299, "  p1 = AIoZ.Instr("" "", p0)"
  Print #299, "  AIoZ.Cut (p0 - 2)" '(p0 + Len(Such) - 1)"
  Print #299, "  AIoZ.Append Mid$(Ursp, p1)"
  Print #299, " END IF"
  Print #299, "End FUNCTION ' AIoZ(Ursp$) AS CString"
  
  
 End If 'Me.nurSchreiben <> 0 THEN
 If Me.nurSchreiben <> 0 Then
  Close #299
 End If
 Close #256
 If Me.nurSchreiben = 0 Then Close #302
 If Not obohnefertig Then
  MsgBox "Fertig"
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
' Feld schon vorhanden
'Dim fehlerrunde%
'If Err.Number = -2147217900 AND fehlerrunde < 3 THEN
' Dim zwicnZ$
' zwicnZ = cnz
' cnz.Close
' DoEvents
' cnz.Open zwicnZ
' Resume
' fehlerrunde = fehlerrunde + 1
'Else
' fehlerrunde = 0
'END IF
If FNr = 1 Then
 MsgBox "Fehler bei dbKopier in '" & Me.SchreibenAuf & "_prot.txt" & "':"
Else
 MsgBox "Fehler bei dbKopier in '" & uVerz & "dbkopier.txt':"
End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dbKopier/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dbKopier

Function datformZ(DaT, obMySQL%) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 Dim obkurz%
 If IsNull(DaT) Then
  datformZ = "null"
 ElseIf (obMySQL <> 0) Then
  datformZ = "'" + Format$(DaT, "yyyy-mm-dd hh:mm:ss") + "'"
 Else
  If VarType(DaT) = vbString Then
   If Len(DaT) < 10 Then
    obkurz = True
   End If
  Else
   If DaT - Int(DaT) = 0 Then
    obkurz = True
   End If
  End If
  If obkurz Then
   datformZ = "#" + Format$(DaT, "mm\/dd\/yy") + "#"
  Else
   datformZ = "#" + Format$(DaT, "mm\/dd\/yy hh:mm:ss") + "#"
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datFormZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datFormZ

Function schl$(keyname$, obZMySQL%)
 If obZMySQL And keyname = "PRIMARY" Then
  schl = keyname & " KEY "
 Else
  schl = "INDEX `" & keyname & "`"
 End If
End Function ' schl

Function JetTyp$(Typ%, Optional size&, Optional obauto%)
 On Error GoTo fehler
 Select Case Typ
   Case 1, 11, 17: JetTyp = "BIT" ' dbboolean
   Case 2: JetTyp = "SMALLINT" ' dbbyte
   Case 3, 4, 19, 131: JetTyp = "INTEGER" ' dbInteger ' dblong
    If obauto Then ' 16
     JetTyp = "AUTOINCREMENT"
    Else
     JetTyp = "INTEGER"
    End If
   Case 5: JetTyp = "double" ' "DECIMAL(15,4)" ' dbcurrency
   Case 6: JetTyp = "FLOAT" 'double(10) ' dbSingle
   Case 7: JetTyp = "DATETIME" ' double(20) ' dbDouble
   Case 133, 134, 135: JetTyp = "DATETIME" ' dbDate
'   Case 9 ' dbbinary
   Case 8, 10, 129, 200, 202:
    If size < 256 Then
     JetTyp = "VARCHAR(" + CStr(size) + ")" ' CHAR macht attributes = 1 => dbfixedfield, mit Leerzeichen am Schluß
    Else
     JetTyp = "LONGTEXT"
    End If
'   Case 11: JetTyp = "VARBINARY(255)"
   Case 12, 201, 203: JetTyp = "LONGTEXT"

   Case 15: JetTyp = "BINARY(16)"
'   Case 16 ' dbbigint
'   Case 18 ' dbchar
   Case 20: JetTyp = "DECIMAL(15,4)" ' dbdecimal"
   Case 23, 135: JetTyp = "TIMESTAMP"
   Case 205: JetTyp = "LONGBINARY"
   Case Else: JetTyp = "?"
 End Select
 If Typ = 10 Then If InStr(JetTyp, "NULL") = 0 Then JetTyp = JetTyp + " NULL"

 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in JetTyp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' JetTyp

Function MySqlTyp$(Typ%, Optional size&, Optional obauto%, Optional obQuNichtMySQL%)
 On Error GoTo fehler
 Select Case Typ
   Case 1, 11, 17: MySqlTyp = "TINYINT(1) UNSIGNED" 'MySqlTyp = "BOOL" ' dbboolean
   Case 16: MySqlTyp = "TINYINT(1)"
   Case 2: MySqlTyp = "SMALLINT" ' dbbyte
   Case 3, 4, 131 ' dblong
    MySqlTyp = "INT(10)"
   Case 5:
    If obQuNichtMySQL Then
     MySqlTyp = "DOUBLE"
    Else
'     Stop
     MySqlTyp = "DOUBLE"
'     MySqlTyp = "DATETIME" ' "DECIMAL(15,4)" ' dbcurrency
    End If
   Case 6: MySqlTyp = "FLOAT" 'double(10) ' dbSingle
   Case 7:
     If obQuNichtMySQL Then
       MySqlTyp = "DATETIME"
     Else
       MySqlTyp = "DOUBLE" ' double(20) ' dbDouble
     End If
   Case 133: MySqlTyp = "DATE"
   Case 134: MySqlTyp = "TIME"
   Case 135: MySqlTyp = "DATETIME" ' dbDate
'   Case 9 ' dbbinary
   Case 8, 10, 129, 200, 202: MySqlTyp = "VARCHAR(" + CStr(size) + ")" ' CHAR macht attributes = 1 => dbfixedfield, mit Leerzeichen am Schluß
   Case 11: MySqlTyp = "VARBINARY(255)"
   Case 12, 201, 203: MySqlTyp = "LONGTEXT"

   Case 15: MySqlTyp = "BINARY(16)"
'   Case 16 ' dbbigint
'   Case 18 ' dbchar
   Case 19: MySqlTyp = "INT(2) UNSIGNED"
   Case 20: MySqlTyp = "DECIMAL(15,4)" ' dbdecimal"
   Case 23: MySqlTyp = "TIMESTAMP"
   Case 205: MySqlTyp = "LONGBINARY"
   Case Else: MySqlTyp = "?"
 End Select
' IF InStr(MySqlTyp, "NULL") = 0 THEN MySqlTyp = MySqlTyp + " NULL"
 If obauto Then ' 16
  MySqlTyp = MySqlTyp & " " & "NOT NULL KEY AUTO_INCREMENT"
 End If

 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MySQLTyp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MySqlTyp

