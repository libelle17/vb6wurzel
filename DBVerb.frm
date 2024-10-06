VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DBVerb 
   Caption         =   "Datenbank-Verbindung herstellen"
   ClientHeight    =   12600
   ClientLeft      =   225
   ClientTop       =   285
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "DBVerb"
   ScaleHeight     =   12600
   ScaleWidth      =   15120
   Begin VB.CheckBox obFilter 
      Caption         =   "&Nur Datenbanken mit den Tabellen:"
      Height          =   255
      Left            =   5160
      TabIndex        =   52
      Top             =   1680
      Width           =   7215
   End
   Begin VB.TextBox Ausgabe 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   51
      Top             =   11040
      Width           =   16095
   End
   Begin VB.TextBox pwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1275
      Width           =   2055
   End
   Begin VB.ComboBox uid 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   915
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox DBKennw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   50
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Tabellen 
      BackColor       =   &H00E7FEE2&
      Height          =   10695
      Left            =   12480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   48
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox Datei 
      Height          =   285
      Left            =   5520
      TabIndex        =   46
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Programmende 
      Caption         =   "&Programmende"
      Height          =   375
      Left            =   11040
      TabIndex        =   44
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0000
      Height          =   375
      Index           =   27
      Left            =   120
      TabIndex        =   43
      Top             =   10560
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "67108864: Enables support for batched statements. This option was enabled in Connector/ODBC 3.51.18."
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   42
      Top             =   10320
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0111
      Height          =   495
      Index           =   25
      Left            =   120
      TabIndex        =   41
      Top             =   9840
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0246
      Height          =   495
      Index           =   24
      Left            =   120
      TabIndex        =   40
      Top             =   9360
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":036C
      Height          =   495
      Index           =   23
      Left            =   120
      TabIndex        =   39
      Top             =   8880
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0497
      Height          =   615
      Index           =   22
      Left            =   120
      TabIndex        =   38
      Top             =   8280
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":05EE
      Height          =   495
      Index           =   21
      Left            =   120
      TabIndex        =   37
      Top             =   7800
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":06DF
      Height          =   495
      Index           =   20
      Left            =   120
      TabIndex        =   36
      Top             =   7320
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "524288: Enable query logging to c:\myodbc.sql(/tmp/myodbc.sql) file. (Enabled only in debug mode.)"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   35
      Top             =   7080
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "262144: Disable transactions."
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   34
      Top             =   6840
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "131072: Add some extra safety checks."
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   33
      Top             =   6600
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "65536: Read parameters from the [client] and [odbc] groups from my.cnf."
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   6360
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":07EE
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "16384: Change BIGINT columns to INT columns (some applications can't handle BIGINT)."
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "8192: Connect with named pipes to a mysqld server running on NT."
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   29
      Top             =   5640
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0886
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2048&: Use the compressed client/server protocol."
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1&024: SQLDescribeCol() returns fully qualified column names."
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&512: Pad CHAR columns to full column length."
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "256: Disable the use of extended fetch (experimental)."
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "128: Force use of ODBC manager cursors (experimental)."
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "64: Ignore use of database name in db_name.tbl_name.col_name."
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&32: Enable or disable the dynamic cursor support. "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1&6: Don't prompt for questions even if driver would like to prompt."
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&8: Don't set any packet limit for results and parameters."
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&4: Make a debug log in C:\myodbc.log on Windows, or /tmp/myodbc.log on Unix variants."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   $"DBVerb.frx":0910
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   12175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&1: The client can't handle that Connector/ODBC returns the real width of a column. This option was removed in 3.51.18."
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   12175
   End
   Begin VB.ComboBox ODBC 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox Benutzer 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   915
      Width           =   2055
   End
   Begin VB.ComboBox DaBa 
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   1635
      Width           =   2415
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton OK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   10440
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Paßwort 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1275
      Width           =   2055
   End
   Begin VB.CommandButton NurLauf 
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox Cpt 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   555
      Width           =   2895
   End
   Begin VB.Label DBKennwLab 
      Caption         =   "Datenbank-&Kennwort:"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label TabellenLab 
      Caption         =   "Tabellen:"
      Height          =   255
      Left            =   12480
      TabIndex        =   47
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Constr 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   12375
   End
   Begin VB.Label ODBCLab 
      Caption         =   "OD&BC-Treiber:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5280
      Top             =   960
      Width           =   735
   End
   Begin VB.Label DabaLab 
      Caption         =   "&Datenbank:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label PaßLab 
      Caption         =   "P&aßwort:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label BenuLab 
      Caption         =   "Ben&utzer:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label CptLabel 
      Caption         =   "Name des Datenbankserver-&PCs:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Image1Ersatz 
      Height          =   255
      Left            =   5880
      TabIndex        =   45
      Top             =   1680
      Width           =   4095
   End
End
Attribute VB_Name = "DBVerb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' erfordert CString.cls
Option Explicit
#Const obPraxis = False
#If obPraxis Then
'Const uVerz$ = "u:\"
#Else
'Const uVerz$ = "c:\u\"
#End If
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal HostName$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public ausaCStr%
Public obQuelle% ' ob aCStr mit quelle o.ä. aufgerufen wurde
Dim FNr&

Private Type HostDeType
 hname As Long
 haliases As Long
 haddrtype As Integer
 hlength As Integer
 haddrlist As Long
End Type

Const Ü1$ = "Datenbank-Verbindung herstellen"
Const NL0$ = "Nur Computer mit MySQL auf&listen", NL1$ = "Alle Computer auf&listen"
Const RegWurzel$ = "Software\GSProducts\"
Const CptLänge% = 15
Public CoStr$, CnStr$
'Dim rTs As New ADODB.Connection
'Dim rTo As New ADODB.Connection
Dim Cpts As New Collection, CptN As New Collection
Dim obNurLauf%
Dim Comp ' Laufvariable für Computer
Dim lCpt$, lBenutzer$, lPaßwort$, luid$, lpwd$, lODBC$, lDaBa$, altCpt$, altODBC$
Public changeStill% ' Veränderungen still vornehmen => keine Folgeereignisse
Dim opt&
Public RegPos$
Public Ü2$ ' soll DBVerb innerhalb eines Programms mit verschiedenen Inhalten gefüllt werden, dann kann hier Unterscheidung getroffen werden (für Überschrift und Registry)
Private BedTbl$() ' Tabellen, die in der Datenbank vorhanden sein müssen
Public wCn As New ADODB.Connection
Attribute wCn.VB_VarHelpID = -1
Public Event wCnAendern(CnStr$)
Dim zuRaisen%
Dim altAusgabe As New CString
Private obAbbruch%

Public Sub rücksetzBedTbl()
 ReDim BedTbl(0)
End Sub ' rücksetzBedTbl()

Public Sub setzBedTbl(Wert$)
 ReDim Preserve BedTbl(UBound(BedTbl) + 1)
 BedTbl(UBound(BedTbl)) = Wert
End Sub ' setzBedTbl(Wert$)

Function HostByName$(name$, Optional x% = 0)
Dim MemIp() As Byte
Dim Y%
Dim HostDeAddress&, HostIp&
Dim IpAddress$
Dim Host As HostDeType
HostDeAddress = gethostbyname(name)
If HostDeAddress = 0 Then
 HostByName = vNS
 Exit Function
End If

Call RtlMoveMemory(Host, HostDeAddress, LenB(Host))
For Y = 0 To x
 Call RtlMoveMemory(HostIp, Host.haddrlist + 4 * Y, 4)
 If HostIp = 0 Then
  HostByName = vNS
  Exit Function
 End If
Next Y
ReDim MemIp(1 To Host.hlength)
Call RtlMoveMemory(MemIp(1), HostIp, Host.hlength)
IpAddress = vNS
For Y = 1 To Host.hlength
 IpAddress = IpAddress & MemIp(Y) & "."
Next Y
HostByName = Left$(IpAddress, Len(IpAddress) - 1)
End Function ' HostByName

Function ShowAllDomains(Optional obneu%) As Collection
  Dim oNameSpace  As Object
  Dim oDomain     As Object
  Dim zuprüfen%
  On Error GoTo fehler
  If obneu Or CptN Is Nothing Then zuprüfen = True Else If CptN.COUNT = 0 Then zuprüfen = True
  If zuprüfen Then
   Set oNameSpace = GetObject("WinNT:")
   For Each oDomain In oNameSpace
'    Debug.Print oDomain.Name
    Call ShowAllComputers(oDomain.name)
   Next
  End If
  On Error GoTo fehler
  Set ShowAllDomains = CptN
  Exit Function
fehler:
  ' vermutlich ist kein WMI installiert
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ShowAllDomains/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ShowAllDomains

Public Sub ShowAllComputers(ByVal strDomain$)
  Dim PrimDomainContr     As Object
  Dim oComputer           As Object
  On Error GoTo fehler
  Set PrimDomainContr = GetObject("WinNT://" & strDomain)
  PrimDomainContr.Filter = Array("Computer")
  On Error Resume Next
  
  For Each oComputer In PrimDomainContr
    Cpts.Add Left$(oComputer.name & Space$(CptLänge), CptLänge) & "| " & HostByName$(oComputer.name)
    CptN.Add UCase$(oComputer.name)
'    Debug.Print oComputer.Name
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
End Sub ' ShowAllComputers(ByVal strDomain$)

Private Sub Abbruch_Click()
' Unload Me
 Call Form_Unload(Cancel:=True)
 Me.Visible = False
 obAbbruch = True
End Sub ' Abbruch_Click()

Private Sub Check1_Click(Index%)
 If Not changeStill Then Call Verbind
End Sub ' Check1_Click(index%)

Private Sub Benutzer_Change()
 If Not changeStill Then Call Verbind
End Sub ' Benutzer_Change()

Private Sub Benutzer_Click()
 If Not changeStill Then Call Verbind
End Sub ' Benutzer_Click()

Private Sub Datei_Change()
' Call doVerbind
End Sub

Private Sub obFilter_Click()
 Call zeigdatenbanken
End Sub ' obFilter_Click()

Private Sub uid_Change()
 If Not changeStill Then Call Verbind
End Sub ' uid_Change()

Private Sub uid_Click()
 If Not changeStill Then Call Verbind
End Sub ' uid_Click()

Private Sub pwd_Change()
 If Not changeStill Then Call Verbind
End Sub ' pwd_Change()

Private Sub Pwd_Click()
 Static altPwd$
 If pwd <> altPwd Then If Not changeStill Then Call Verbind
 altPwd = pwd
End Sub ' Pwd_Click()

Private Sub Datei_Click()
 Dim fileflags As FileOpenConstants
 Dim filefilter$
 CommonDialog1.DialogTitle = "Datenbankdatei auswählen:"
 'Set the default file name AND filter
 CommonDialog1.initDir = uVerz
 CommonDialog1.Filename = vNS
 filefilter = "Access-Dateien(*.mdb)|*.mdb|Alle Dateien (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 CommonDialog1.flags = fileflags
 CommonDialog1.ShowOpen
 Me.Datei = CommonDialog1.Filename
 Call Me.Verbind
End Sub ' Datei_Click()

Private Sub DBKennw_Change()
  If Not changeStill Then Call Verbind
End Sub ' DBKennw_Change()

Private Sub DBKennw_Click()
 Static altDBKennw$
 If Me.DBKennw <> altDBKennw Then If Not changeStill Then Call Verbind
 altDBKennw = DBKennw
End Sub ' DBKennw_Click

Private Sub DBKennw_GotFocus()
 Me.DBKennw.SelStart = 0
 Me.DBKennw.SelLength = Len(Me.DBKennw)
End Sub ' DBKennw_GotFocus()

Private Sub Form_Deactivate()
 Ü2 = vNS
 ReDim BedTbl(0)
End Sub ' Form_Deactivate()

Private Sub Form_GotFocus()
'
End Sub

Private Sub Form_Initialize()
 ReDim BedTbl(0)
End Sub ' Form_Initialize()

Private Sub Form_LinkOpen(Cancel As Integer)
'
End Sub

Private Sub Form_Paint()
'
End Sub

Private Sub Paßwort_Change()
  If Not changeStill Then Call Verbind
End Sub ' Paßwort_Change()

Private Sub Paßwort_Click()
 Static altpaßwort$
 If Paßwort <> altpaßwort Then If Not changeStill Then Call Verbind
 altpaßwort = Paßwort
End Sub ' Paßwort_Click

Private Sub daba_Change()
 If Not changeStill Then
'  Me.obQuelle = True
  Call Verbind
'  Me.obQuelle = False
 End If
End Sub ' daba_Change()

Private Sub daba_Click()
 If Not changeStill Then
'  Me.obQuelle = True
  Call Verbind
'  Me.obQuelle = False
 End If
End Sub ' daba_Click()

Private Sub odbc_Change()
' IF Not changeStill THEN
'  Call Verbind
  If Not ausaCStr Then
   Call odbcAngl
  End If
' END IF
End Sub ' odbc_Change()

Private Sub odbcAngl()
 Dim i%
 Dim altChangeStill%
 Me.DBKennw.Visible = False
 Me.DBKennwLab.Visible = False
 If InStrB(Me.ODBC, "Acc") > 0 Then
  Me.Datei.Visible = True
  For i = 0 To Me.Check1.COUNT - 1
   Me.Check1(i).Visible = False
  Next i
  Me.DaBa.Visible = False
  Me.DabaLab.Visible = False
  altChangeStill = changeStill
  changeStill = True
  If LenB(Me.Benutzer) = 0 Then
   Me.Benutzer = "admin"
   Me.Paßwort = vNS
  End If
  changeStill = altChangeStill
'  Me.Paßwort = vns
  Me.uid.Visible = False
  Me.pwd.Visible = False
  Me.Paßwort.Visible = True
  Me.Benutzer.Visible = True
  Me.DBKennw.Visible = True
  Me.DBKennwLab.Visible = True
  Me.Cpt.Visible = False
  Me.CptLabel.Visible = False
  Me.NurLauf.Visible = False
 ElseIf InStrB(Me.ODBC, "MySQL") <> 0 Or InStr(Me.ODBC, "MSDASQL") > 0 Then
  Me.Datei.Visible = False
  For i = 0 To Me.Check1.COUNT - 1
   Me.Check1(i).Visible = True
  Next i
  Me.DaBa.Visible = True
  Me.DabaLab.Visible = True
  altChangeStill = changeStill
  changeStill = True
  If LenB(Me.uid) = 0 Then
   Me.uid = "mysql"
   Me.pwd = vNS
  End If
  changeStill = altChangeStill
  Me.uid.Visible = True
  Me.pwd.Visible = True
  Me.Paßwort.Visible = False
  Me.Benutzer.Visible = False
  Me.Cpt.Visible = True
  Me.CptLabel.Visible = True
  Me.NurLauf.Visible = True
 End If
 If Not changeStill Then Call doVerbind
End Sub ' odbcAngl

Private Sub ODBC_Click()
 If Not changeStill Then
  Call odbcAngl
  Call Verbind ' ' Kommentar testweise entfernt und Befehl nachgestellt 20.9.08
 End If
End Sub 'odbc_Click

Private Sub ODBC_DropDown()
  If Me.ODBC.ListCount = 0 Then
   Call listOdbc
  End If
End Sub ' ODBC_DropDown()

Private Sub Cpt_Change()
  If Not changeStill Then Call Verbind
End Sub ' Cpt_Change()

Private Sub Cpt_GotFocus()
'
End Sub

Private Sub Cpt_Click()
 If Not changeStill Then Call Verbind
End Sub ' Cpt_Click()

Private Sub Cpt_DropDown()
' Call Verbind
  If Me.Cpt.ListCount = 0 Then
   Screen.MousePointer = vbHourglass
   Call ShowAllDomains
   Call CptListeGanz
   Screen.MousePointer = vbNormal
  End If
End Sub ' Cpt_DropDown()

Private Sub Cpt_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 115 And Shift = 0 Then ' F4
 End If
'  Call Verbind
End Sub ' Cpt_KeyDown(

Private Sub Cpt_LostFocus()
' Call Verbind
End Sub ' Cpt_LostFocus

Private Sub Cpt_Scroll()
' Call Verbind
End Sub ' Cpt_Scroll

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo fehler
 Select Case KeyCode
  Case 27
'  Unload Me
   Call Form_Unload(Cancel:=True)
   Me.Visible = False
  Case 13
   Call OK_Click
  Case Else
 End Select
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_KeyDown/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_KeyDown(

Private Sub NurLauf_Click()
 Dim rTs As New ADODB.Connection
' dim rs As New ADODB.Recordset ' geht auch nicht schneller
 Dim i&, otrei$, odb$
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 If obNurLauf Then
  Call CptListeGanz
 Else
   For i = 0 To Me.ODBC.ListCount
    odb = Me.ODBC.List(i)
    If InStrB(UCase$(odb), "MYSQL") <> 0 Or InStr(1, odb, "MSDASQL", vbTextCompare) > 0 Then
     otrei = odb
     Exit For
    End If
   Next i
'  For Each Comp in Cpts
  For i = Me.Cpt.ListCount - 1 To 0 Step -1
   Set rTs = Nothing
'   SET rs = Nothing
   Err.Clear
   On Error Resume Next
   rTs.Open "DRIVER={" & Me.ODBC & "};server=" & Trim$(Left$(Me.Cpt.List(i), 15)) & ";option=3;uid=" & Me.uid & ";pwd=" & Me.pwd & ";"
'   rs.Open "SELECT * FROM mysql.user LIMIT 1", "DRIVER={" & Me.ODBC & "};server=" & trim$(LEFT(Me.Cpt.List(i), 15)) & ";option=3;uid=" & Me.uid & ";pwd=" & Me.pwd & ";", adOpenStatic, adLockReadOnly
   If Err.Number = 0 Or InStrB(Err.Description, "denied") <> 0 Then ' Access denied
    On Error GoTo fehler
'    Me.Cpt.AddItem Me.Cpt.List(i)
   Else
    Me.Cpt.RemoveItem i
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

Private Sub CptListeGanz()
 Dim i%
 On Error GoTo fehler
' Me.Cpt.Width = 3500
'  altCpt = Me.Cpt
'  Me.Cpt.Clear
'  changeStill = True
'  Me.Cpt = altCpt
'  changeStill = False
 Dim ischoda%
 For i = Me.Cpt.ListCount - 1 To 0 Step -1
  If Me.Cpt.List(i) = Me.Cpt.Text And Not ischoda Then
   ischoda = True
  Else
   Me.Cpt.RemoveItem i
  End If
 Next i
 For Each Comp In Cpts
'  IF Comp <> Me.Cpt.Text THEN
   Me.Cpt.AddItem Comp
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
End Sub ' CptListeGanz()

Private Sub Form_Load()
 Dim i%
 On Error GoTo fehler
 Me.Datei.Visible = False
 Me.DBKennw.Visible = False
 Me.DBKennwLab.Visible = False
' For i = 0 To Me.Check1.Count - 1
'  Me.Check1(i).Visible = True
' Next i
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Load/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_Load

Private Sub Form_Activate()
 On Error GoTo fehler
 Screen.MousePointer = vbHourglass
 Call Verbind
 Screen.MousePointer = vbNormal
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Activate/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_Activate()

' aufgerufen in Form_Load
Private Sub RegLaden(Ü$, Optional nuranfangs%)
 Dim neuS$, neuB&
 Static angefangen%, altÜ$
 Dim cR As New Registry
 If Not nuranfangs Or Not angefangen Or altÜ <> Ü Then
  On Error Resume Next
  changeStill = True
  RegPos = RegWurzel & App.EXEName & "\DBVerb"
  If LenB(Ü2) <> 0 Then RegPos = RegPos & "\" & Ü
  neuS = cR.ReadKey("ODBC", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.ODBC = neuS
  neuS = cR.ReadKey("Paßwort", RegPos, HKEY_CURRENT_USER)
'  IF neuS <> vns THEN ' Kommentar 15.8.09 wg. Faxdopp, für Access Paßwort leer
   Me.Paßwort = neuS
'  END IF
  neuS = cR.ReadKey("Datenbank", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.DaBa = neuS Else Me.DaBa = vNS
  neuS = cR.ReadKey("uid", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.uid = neuS
  If LenB(Me.uid) = 0 Then Me.uid = "mysql"
  neuS = cR.ReadKey("pwd", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.pwd = neuS Else If LenB(Me.DaBa) = 0 Then Me.pwd = vNS
  neuS = cR.ReadKey("server", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.Cpt = neuS
  neuS = cR.ReadKey("options", RegPos, HKEY_CURRENT_USER)
  If neuB <> 0 Then opt = neuB
  neuS = cR.ReadKey("Datei", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.Datei = neuS Else Me.Datei = vNS
  neuS = cR.ReadKey("Benutzer", RegPos, HKEY_CURRENT_USER)
  If LenB(neuS) <> 0 Then Me.Benutzer = neuS
  If LenB(Me.Benutzer) = 0 Or LenB(Me.Datei) = 0 Then Me.Benutzer = "admin"
  neuS = cR.ReadKey("DBKennw", RegPos, HKEY_CURRENT_USER)
'  neuS = fWertLesen(HCU, RegPos, "DBKennw")
  If LenB(neuS) <> 0 Then Me.DBKennw = neuS Else If LenB(Me.Datei) = 0 Then Me.DBKennw = vNS
  Call setzeOpt
  changeStill = False
  angefangen = True
  altÜ = Ü
 End If ' not nuranfangs OR not angefangen
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegLaden/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' regladen

Public Sub RegSpeichern()
 Dim cR As New Registry
 On Error GoTo fehler
  cR.WriteKey Me.ODBC, "ODBC", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.Cpt, "Server", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.Benutzer, "Benutzer", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.pwd, "pwd", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.DaBa, "Datenbank", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey opt, "options", RegPos, HKEY_CURRENT_USER, REG_DWORD
  cR.WriteKey Me.uid, "uid", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.Paßwort, "Paßwort", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.Datei, "Datei", RegPos, HKEY_CURRENT_USER, REG_SZ
  cR.WriteKey Me.DBKennw, "DBKennw", RegPos, HKEY_CURRENT_USER, REG_SZ
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in RegSpeichern/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' RegSpeichern

Private Sub Form_Unload(Cancel As Integer)
 On Error GoTo fehler
 If Cancel Then
  Call RegLaden(Ü2)
 Else
  Call RegSpeichern
 End If
 Exit Sub
fehler:
 Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Form_Unload/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Form_Unload(Cancel As Integer)

Sub setzeOpt()
 Dim i%, aktopt&
 aktopt = opt
 changeStill = True
 For i = 0 To Me.Check1.COUNT - 1
  If aktopt Mod 2 = 1 Then
   Me.Check1(i) = 1
   aktopt = aktopt - 1
  End If
  aktopt = aktopt / 2
 Next i
 changeStill = False
End Sub ' setzeOpt()

Sub rechneOpt()
 Dim i%, lauf&
 opt = 0
 lauf = 1
 For i = 0 To Me.Check1.COUNT - 1
  If Me.Check1(i) = 1 Then opt = opt + lauf
  lauf = lauf + lauf
 Next i
End Sub ' rechneOpt()

Function doVerbind%(Optional Tabelle$, Optional ErrDes$)
  Dim s1$, s2$
  Dim rX As New ADOX.Catalog, rxt As ADOX.Table
  Dim ErrDescr$
  ' CnStr = ggf. mit Datenbank, CoStr = ohne, ConStr = mit gesterntem Paßwort
  On Error Resume Next
    If LenB(Me.ODBC) = 0 Then
     If LenB(Me.CnStr) <> 0 Then
      Dim p1&, p2&
      p1 = InStr(1, Me.CnStr, "DRIVER=", vbTextCompare)
      If p1 <> 0 Then
       p2 = InStr(p1, Me.CnStr, ";")
       If p2 <> 0 Then
        Me.ODBC = Mid$(Me.CnStr, p1 + 8, p2 - p1 - 9)
        Me.CnStr = vNS
       End If
      End If
      
     End If
    End If
    If InStrB(Me.ODBC, "MySQL") <> 0 Or InStr(Me.ODBC, "MSDASQL") > 0 Then
     CnStr = "DRIVER={" & Me.ODBC & "};server=" & Trim$(Left$(Me.Cpt, CptLänge)) & ";"
     Call rechneOpt
     CnStr = CnStr & "option=" & opt & ";"
     CoStr = CnStr
     If LenB(Me.DaBa) <> 0 Then
      CnStr = CnStr & "database=" & Me.DaBa & ";"
     End If
     CnStr = CnStr & "uid=" & Me.uid & ";pwd="
     CoStr = CoStr & "uid=" & Me.uid & ";pwd="
     Constr = CnStr
     CnStr = CnStr & Me.pwd & ";"
     CoStr = CoStr & Me.pwd & ";"
     Constr = Constr & String$(Len(Me.pwd), "*") & ";"
    ElseIf InStrB(Me.ODBC, "Acc") <> 0 Or InStrB(Me.ODBC, ".Jet.") <> 0 Then
     CnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Me.Datei & "';"
     Constr = CnStr
     If LenB(Me.DBKennw) <> 0 Then
      CnStr = CnStr & "Jet OLEDB:Database Password="
      Constr = CnStr & String$(Len(Me.Paßwort), "*") & ";"
      CnStr = CnStr & Me.DBKennw & ";"
     End If
     CnStr = CnStr & "user id=" & Me.Benutzer & ";password="
     Constr = Constr & "user id=" & Me.Benutzer & ";password="
     CoStr = CnStr
     CnStr = CnStr & Me.Paßwort & ";"
     CoStr = CoStr & Me.Paßwort & ";"
     Constr = Constr & String$(Len(Me.Paßwort), "*") & ";"
    End If
'    CnStr = CnStr & "user id=" & Me.Benutzer & ";password=" ' geht auch, aber wer weiß, ob's nicht Fehler macht
'    Constr = Constr & "user id=" & Me.Benutzer & ";password="
    
'    zuRaisen = True
'    IF Not wCn Is Nothing THEN
'     IF wCn.ConnectionString = Me.CnStr THEN
      zuRaisen = False
'     END IF
'    END IF
again:
    Err.Clear
    Set wCn = Nothing
    wCn.Open Me.CnStr
    DBCnS = Me.CnStr
'    DefaultDatabase = wCn.DefaultDatabase
    If Err.Number = 0 Then
'     GoTo again:
     doVerbind = 0 ' verbunden
     If LenB(Tabelle) <> 0 Then
      Err.Clear
      If InStrB(Me.ODBC, "MySQL") <> 0 Or InStr(Me.ODBC, "MSDASQL") > 0 Then
       myEFrag "SELECT * FROM `" & Tabelle & "` LIMIT 1", , wCn
'       Call myEFrag("SELECT * FROM `" & Tabelle & "` LIMIT 1",,wCn)
      ElseIf InStrB(Me.ODBC, "Acc") <> 0 Then
       Call wCn.Execute("SELECT top 1 * FROM `" & Tabelle & "`")
      Else
       rX.ActiveConnection = CnStr ' wCn.ConnectionString
       Set rxt = rX.Tables(Tabelle)
      End If
      If Err.Number <> 0 Then
       ErrDes = Err.Description
       Me.DaBa = vNS
       doVerbind = 1
      End If
     End If
    Else
 '    GoTo again:
     ErrDescr = Err.Description
     If InStrB(ErrDescr, "Unknown database") <> 0 Then
      ErrDes = Err.Description
      Err.Clear
      Set wCn = Nothing
      Me.CnStr = Me.CoStr
      wCn.Open Me.CnStr
      DBCnS = Me.CnStr
     End If
     If Err.Number = 0 Then
      ErrDes = Err.Description
      doVerbind = 1 ' verbunden, Datenbank nicht gefunden
     Else
      ErrDes = Err.Description
      doVerbind = 2 ' nicht verbindbar
'      Call Shell(App.Path + "\..\nachricht\nachricht.exe " & App.EXEName & " Verbindung gescheitert mit: " & Me.CnStr)
      Me.Ausgeb "Verbindung gescheitert mit: " & Me.CnStr & ": " & ErrDes, True
'      GoTo again
     End If
    End If
   On Error Resume Next
   If zuRaisen Then RaiseEvent wCnAendern(Me.CnStr)
End Function ' doVerbind

Sub zeigdatenbanken()
 Dim rSch As New ADODB.Recordset, rs1 As New ADODB.Recordset, obMySQL%, i&, j&
 Dim tTbl%() ' teste Tabellen: wenn true, ist die jeweilige BedTbl enthalten
 Static altwcn$, altobfilter%
 On Error GoTo fehler
 If Not wCn Is Nothing Then
  If CnStr <> altwcn Or Me.obFilter <> altobfilter Then ' wCn.ConnectionString
   On Error Resume Next
   Set rSch = wCn.OpenSchema(adSchemaCatalogs) '-> bei Access Fehler, manchmal auch bei MySQL
'   IF Err.Number <> 0 THEN
'    SET rSch = wCn.Execute("SHOW DATABASES") ' geht dann auch nicht
'   END IF
   If Err.Number = 0 Then
    On Error GoTo fehler
' alle Listeneinträge außer dem aktuell im Textfeld stehenden entfernen
    Dim ischoda%
    For i = Me.DaBa.ListCount - 1 To 0 Step -1
     If Me.DaBa.List(i) = Me.DaBa.Text And Not ischoda Then
      ischoda = True
     Else
      Me.DaBa.RemoveItem i
     End If
    Next i
' jetzt Listeneinträge neu erstellen
    If UBound(BedTbl) <> 0 Then ReDim tTbl(UBound(BedTbl))
    Dim obVgl%
    If Me.obFilter = 0 Then
     obVgl = 0
    Else
     For i = 1 To UBound(BedTbl)
      If LenB(BedTbl(i)) <> 0 Then obVgl = True: Exit For
     Next i
    End If
    Do While Not rSch.EOF
     Dim obenthalten%
     obenthalten = True
' Wenn Bedingungstabellen angegeben
     If obVgl Then
'      IF wCn.ConnectionString = "Provider=MSDASQL.1;" THEN
'       obMySQL = (wCn.Properties("DBMS Name") = "MySQL")
'      Else
       obMySQL = (InStrB(Me.CnStr, "MySQL") <> 0 Or InStr(Me.CnStr, "MSDASQL") <> 0)  ' wCn.ConnectionString
'      END IF

      If obMySQL Then
       Set rs1 = myEFrag("SHOW TABLES FROM `" & rSch!catalog_name & "`", , wCn)
      Else
       Set rs1 = wCn.OpenSchema(adSchemaTables, Array(UCase$(rSch!catalog_name), vNS, Empty, "TABLE"))  'Array(, vns & rSch!catalog_name & vns)
      End If
      If Err.Number = 0 Then
       For i = 1 To UBound(BedTbl)
        tTbl(i) = 0
       Next i
       Dim schonalletrue%
       schonalletrue = 0
       If rs1.BOF Then obenthalten = False
       Do While Not rs1.EOF
        Dim Vgl$
        If obMySQL Then Vgl = rs1.Fields(0) Else Vgl = rs1!table_name
        For i = 1 To UBound(BedTbl)
         If UCase$(BedTbl(i)) = UCase$(Vgl) Then
          tTbl(i) = True
          schonalletrue = True
          For j = 1 To UBound(BedTbl)
           If tTbl(i) = 0 Then
            schonalletrue = 0: Exit For
           End If
          Next j
          If schonalletrue Then Exit Do
          Exit For
         End If
        Next i
        rs1.Move 1
       Loop
       If Not schonalletrue Then
        For i = 1 To UBound(BedTbl)
         If tTbl(i) = 0 Then
          obenthalten = False
          Exit For
         End If
        Next i
       End If
      End If ' Err.Number = 0 THEN
     End If ' obvgl
     If obenthalten Then
      Dim schonda%
      schonda = 0
      For i = 0 To Me.DaBa.ListCount
       If Me.DaBa.List(i) = rSch!catalog_name Then schonda = True: Exit For
      Next i
      If Not schonda Then 'If rSch!catalog_name <> Me.DaBa.Text THEN
       Me.DaBa.AddItem rSch!catalog_name
      End If
     End If ' obenthalten
     rSch.Move 1
    Loop ' while not rSch.EOF
   End If ' Err.number = 0
  End If ' wcn.ConnectionString <> altwcn THEN
  altwcn = Me.wCn
  altobfilter = Me.obFilter
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in zeigdatenbanken/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' zeigdatenbanken

Sub zeigtabellen()
 Dim rSch As New ADODB.Recordset, rs1 As New ADODB.Recordset, obMySQL%
 Me.Tabellen = vNS
' IF wCn.ConnectionString = "Provider=MSDASQL.1;" THEN
'  obMySQL = (wCn.Properties("DBMS Name") = "MySQL")
' Else
  obMySQL = (InStrB(Me.CnStr, "MySQL") <> 0 Or InStr(Me.CnStr, "MSDASQL") > 0) ' wCn.ConnectionString ' geht auch mit obMySQL = 0, aber viel langsamer
' END IF
 If obMySQL And CnStr <> Me.CoStr Then ' wCn.ConnectionString
  Set rs1 = myEFrag("SHOW TABLES FROM `" & wCn.Properties("Current Catalog") & "`", , wCn)
  Do While Not rs1.EOF
   Me.Tabellen = Me.Tabellen & rs1.Fields(0) & vbCrLf
   rs1.MoveNext
  Loop
 ElseIf InStrB(Me.ODBC, "Acc") <> 0 Then
  On Error GoTo fehler
  Set rSch = Nothing
  Dim rc As New ADOX.Catalog
  rc.ActiveConnection = wCn
  Dim tbl1
  For Each tbl1 In rc.Tables
   Me.Tabellen = Me.Tabellen & tbl1.name & vbCrLf
  Next tbl1
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in zeigtabellen/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' zeigTabellen

Sub Verbind()
 Static DBChange%
 Dim i&, erg%, altUser$
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
' IF Me.Cpt = vns THEN Call RegLaden
 If UCase$(Me.Cpt) <> UCase$(lCpt) Or UCase$(Me.Benutzer) <> UCase$(lBenutzer) Or Me.Paßwort <> lPaßwort Or UCase$(lODBC) <> UCase$(Me.ODBC) Or UCase$(Me.uid) <> UCase$(luid) Or UCase$(Me.pwd) <> UCase$(lpwd) Then
  Screen.MousePointer = vbHourglass
  On Error Resume Next
  Me.Image1Ersatz = "nicht verbunden"
  Me.Tabellen = vNS
  Me.Image1 = LoadPicture(App.path & "\..\icons\tele\164.ico") '"\..\icons\tele\636I93.ico")
  If Err.Number <> 0 Then
   Me.Image1 = LoadPicture(App.path & "\164.ico")
  End If
  On Error GoTo fehler
  Me.Refresh
'  DoEvents
  On Error Resume Next
'  For Each Comp in Cpts
'   IF ucase$(Me.Cpt) = ucase$(Comp) THEN
    erg = doVerbind
    If erg < 2 Then
     If erg = 1 Then
       Me.Image1Ersatz = "verbunden, Datenbank nicht gefunden"
       Me.Tabellen = vNS
       Me.Image1 = LoadPicture(App.path & "\..\icons\tele\131.ico") '"\..\icons\tele\636I93.ico")
       If Err.Number <> 0 Then
        Me.Image1 = LoadPicture(App.path & "\131.ico")
       End If
     End If
     Call zeigdatenbanken
     Call zeigtabellen
    End If ' erg > 0
    If erg = 0 Then
     On Error Resume Next
     Me.Image1Ersatz = "verbunden"
     Me.Image1 = LoadPicture(App.path & "\..\icons\tele\156.ico") '"\..\icons\tele\636I91.ico")
     If Err.Number <> 0 Then
      Me.Image1 = LoadPicture(App.path & "\156.ico") '"\..\icons\tele\636I91.ico")
     End If
     Set rs = Nothing
     altUser = Me.uid
     Dim ischoda%
'     Me.uid.Clear
     For i = Me.uid.ListCount - 1 To 0 Step -1
      If Me.uid.List(i) = Me.uid.Text And Not ischoda Then
       ischoda = True
      Else
       Me.uid.RemoveItem i
      End If
     Next i
'     changeStill = True
     On Error Resume Next
     myFrag rs, "SELECT DISTINCT user FROM mysql.`user`;", adOpenStatic, wCn, adLockReadOnly
     If Not rs.BOF Then
      Do While Not rs.EOF
       If Me.uid.Text <> rs.Fields(0) Then
        Me.uid.AddItem rs.Fields(0)
        If rs.Fields(0) = altUser Then
         Me.uid.ListIndex = Me.uid.ListCount - 1
        End If
       End If
       rs.Move 1
      Loop
     End If ' not rs.bof then
'     Me.uid = altUser
'     changeStill = False
     
     On Error GoTo fehler
    End If
'    Exit For
'   END IF
'  Next Comp
  Screen.MousePointer = vbNormal
 End If
 lCpt = Me.Cpt
 lBenutzer = Me.Benutzer
 lPaßwort = Me.Paßwort
 lDaBa = Me.DaBa
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Verbind/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Verbind

Private Sub OK_Click()
' Unload Me
 Call Form_Unload(Cancel:=False)
 Me.Visible = False
End Sub ' ok_click

Public Sub listOdbc()
 Const strComputer$ = "."
 Dim objRegistry As Object, strKeyPath$, altChangeStill%
 Dim i%, runde%, rrunde%, stand%, arrValueNames, arrValueTypes, sV$, strValue$, Sch$(), schz&, Inh$()
 Dim cR As New Registry
 Dim inh0$
 Const HKEY_LOCAL_MACHINE = &H80000002
 On Error GoTo fehler
 For i = Me.ODBC.ListCount - 1 To 0 Step -1
  Me.ODBC.RemoveItem i
 Next i
 
' altODBC = Me.ODBC
' Me.ODBC.Clear
' altChangeStill = changeStill
' changeStill = True
' Me.ODBC = altODBC
' changeStill = altChangeStill
 On Error Resume Next
 Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 If Err.Number <> 0 Then ' wnenn wmi nicht läut
  On Error GoTo fehler
  cR.EnumerateVALUES Sch, schz, Inh, "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", HKEY_LOCAL_MACHINE
'  Call regEnumVal("SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", Sch, Inh)
  For i = 1 To UBound(Sch)
'   inh0 = cR.ReadKey(Sch(i))
   If Inh(i) = "Installed" Then Me.ODBC.AddItem Sch(i)
  Next i
'  Me.ODBC.AddItem "MySQL ODBC 5.1 Driver"
'  Me.ODBC.AddItem "MySQL ODBC 5.1 Driver"
'  Me.ODBC.AddItem "Achtung: WMI-Dienst nicht gestartet; Auswahl könnte nicht den verfügbaren Treibern entsprechen!"
 Else
  On Error GoTo fehler
  strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
  objRegistry.EnumVALUES HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
  stand = 0
  For runde = 0 To 3
   For i = 0 To UBound(arrValueNames)
    sV = arrValueNames(i)
    ' rrunde = richtige Runde für jeden Treiber; zuerst die übrigen nach Sortierung, dann die interessanten nach Stand
    If InStrB(sV, "MySQL ODBC 5") <> 0 Then
     rrunde = 1
    ElseIf InStrB(sV, "Microsoft Access-Treiber") <> 0 Then
     rrunde = 2
    ElseIf InStrB(sV, "Access") <> 0 Then
     rrunde = 3
    Else
     rrunde = 0
    End If
    If runde = rrunde Then
     objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, sV, strValue
'    Debug.Print arrValueNames(i) & " -- " & strValue
     If strValue = "Installed" Then
       If runde = 0 Then
        Me.ODBC.AddItem sV
       Else
        Me.ODBC.AddItem sV, stand
        stand = stand + 1
       End If
     End If
    End If
   Next i
  Next runde
 End If
 If Me.ODBC.ListCount > 0 And LenB(Me.ODBC) = 0 Then
  Me.ODBC = Me.ODBC.List(0)
 End If
 Exit Sub
fehler:
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in listODBC/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' listodbc

Private Sub Paßwort_GotFocus()
 Me.Paßwort.SelStart = 0
 Me.Paßwort.SelLength = Len(Me.Paßwort)
End Sub ' paßwort_gotfocus

Private Sub Programmende_Click()
 ProgEnde
End Sub ' Programmende_Click

Private Function Überschrift(Ü$)
 Me.Ü2 = Ü
 Me.Caption = App.Title & IIf(LenB(Ü2) <> 0, " (" & Ü2 & ")", vNS) & ": " & Ü1
' Call Verbind
End Function ' Überschrift

Sub gemeinsam(TBName$, Ü$)
   If TBName <> "--multi" Then
    Call Me.rücksetzBedTbl
    Call Me.setzBedTbl(TBName)
   End If
   If LenB(TBName) = 0 Then
    Me.obFilter.Caption = "Alle Datenbanken anzeigen": Me.obFilter = 0: Me.obFilter.Enabled = False
   Else
    Me.obFilter.Caption = "&Nur Datenbanken anzeigen mit der Tabelle: " & TBName: Me.obFilter = 1
   End If
   Call Überschrift(Ü)
End Sub ' gemeinsam

Function cnVorb$(DBName$, TBName$, Optional Ü$, Optional obregneu%, Optional RegNichtLaden%)
  Dim ErrNumber&, altChangeStill%
  Dim cR As New Registry
  Static innen%
  On Error GoTo fehler
  If Not innen Then
   Call gemeinsam(TBName, Ü)
   If Not RegNichtLaden Then Call RegLaden(Ü, nuranfangs:=IIf(obregneu, False, True))
   If LenB(DBName) <> 0 Then
    altChangeStill = changeStill
    changeStill = True
    Me.DaBa = DBName
    changeStill = altChangeStill
   End If
   Do
    ErrNumber = doVerbind(TBName)
    If (TBName = "--multi" And ErrNumber = 1) Or ErrNumber = 0 Then Exit Do
    On Error GoTo fehler
    innen = True
    On Error Resume Next ' Modulares Formular kann in diesem Zusammenhang nicht gezeigt werden
    Me.Show 1
    If Err.Number <> 0 Then Exit Function
    On Error GoTo fehler
    innen = False
   Loop
'  IF RegPos = vns THEN ' beim ersten Mal im Programm
'   Me.DaBa = DBName ' hier Load me enthalten
'  Else
'   Me.DaBa = DBName
'   Call Form_Load
'  END IF
   cnVorb = Me.CnStr ' .wCn.ConnectionString ' 28.12.08
   If LCase$(Left$(Ü, 5)) = "admin" Then
    cR.WriteKey vNS, "Paßwort", RegPos, HKEY_CURRENT_USER, REG_SZ
'    Call fStSpei(HCU, RegPos, "Paßwort", vns)
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
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cnVorb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' cnVorb

Function Auswahl$(DBName$, TBName$, Optional Ü$)
 Dim altCn$
 On Error Resume Next
 altCn = Me.CnStr ' .wCn.ConnectionString ' 28.12.08
 On Error GoTo fehler
 Call gemeinsam(TBName, Ü)
 Call RegLaden(Ü, nuranfangs:=True)
 If LenB(DBName) <> 0 Then Me.DaBa = DBName ' hier Load me enthalten
 Call doVerbind
zeig:
 Me.Show 1
 If obAbbruch Then
  On Error Resume Next
  If wCn.State = 0 Or Me.CnStr <> altCn Then ' .wCn.ConnectionString ' 28.12.08
   wCn.Close
   Me.CnStr = altCn
   wCn.Open Me.CnStr
   DBCnS = Me.CnStr
  End If
  On Error GoTo fehler
  obAbbruch = False
 Else
'  zuRaisen = True
'  IF Not wCn Is Nothing THEN
'   IF wCn.ConnectionString = Me.CnStr THEN
    zuRaisen = False
'   END IF
'  END IF
  Set wCn = Nothing
  wCn.Open Me.CnStr
  DBCnS = Me.CnStr
  If zuRaisen Then RaiseEvent wCnAendern(Me.CnStr)
 End If
 Auswahl = Me.CnStr
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Auswahl/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"):
  If Err.Number = -2147467259 Then
   Resume zeig
  Else
   Resume
  End If
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Auswahl

Public Function Ausgeb(Text$, obDauer%)
 Dim aktText As New CString
 If Not Me.Visible Then
'  Debug.Print Text
 Else
  aktText = Text
  aktText.Append vbCrLf
  aktText.Append altAusgabe
  aktText.Cut 3000
  Me.Ausgabe = aktText
  If obDauer <> 0 Then
   altAusgabe = aktText
  End If
  If InStrB(Text, "READ-COMMITTED") <> 0 Then
   MsgBox "Beinahe-Stop in Ausgeb:" & vbCrLf & "instrb(text, 'READ-COMMITTED') <> 0" & vbCrLf & "Text: " & Text
  End If
  DoEvents
 End If
End Function ' Ausgeb

'Public FUNCTION Ausgeb(Text$, obDauer%)
' IF Not Me.Visible THEN
'  Debug.Print Text
' Else
'  Me.Ausgabe = Text & vbCrLf & altAusgabe.Value
'  IF obDauer THEN
'   altAusgabe = Me.Ausgabe
'   altAusgabe.Cut 7000
'  END IF
'  IF InStrB(Text, "READ-COMMITTED") <> 0 THEN
''       oboffenlassen = 0    ' 27.9.09 für Wien
''       Ende                 ' 27.9.09 für Wien
'   MsgBox "Beinahe-Stop in Ausgeb:" & vbCrLf & "instrb(text, 'READ-COMMITTED') <> 0" & vbCrLf & "Text: " & Text
'  END IF
'  DoEvents
' END IF
'End FUNCTION ' Ausgeb

' 15.10.10: scheint nicht vorzukommen
' 28.6.24: nur in testgetAllDB
Function getAllDB%(Tabl$, acn() As ADODB.Connection, Optional uid$ = "...", Optional pwd$ = "...", Optional obLeere%, Optional acSt)
 Dim Cpt, db$, eintragen%, Stri(1) As New CString, i%, runde%, ErrNr&
 Dim MyCn As ADODB.Connection, rdb As ADODB.Recordset, rHa As New ADODB.Recordset
 ReDim acn(0)
 ReDim acSt(0)
 Call Me.ShowAllDomains
 For runde = 1 To 2
  For Each Cpt In CptN
   If (runde = 1 And Cpt = LiName) Or (runde = 2 And Cpt <> LiName) Then
   Set MyCn = New ADODB.Connection
   On Error Resume Next
   Err.Clear
   Stri(0) = "Provider=MSDASQL.1;Extended Properties=""DRIVER={" & ODBCStr & "};OPTION=3;PWD="
   Stri(1) = Stri(0)
   Stri(0).Append pwd
   Stri(1).Append "***"
   For i = 0 To 1
    Stri(i).AppVar Array(";PORT=0;SERVER=", Cpt, ";UID=", uid, """")
   Next i
   MyCn.Open Stri(0)
   If Err.Number = 0 Then
    If LenB(Tabl) <> 0 Then
     Set rdb = myEFrag("SHOW DATABASES", , MyCn)
    End If
    db = vNS
    Do
     If LenB(Tabl) <> 0 Then
      eintragen = False
      Err.Clear
      myEFrag "USE `" & rdb.Fields(0) & "`", , MyCn, True, ErrNr
      If ErrNr = 0 Then
       Set rHa = Nothing
'       rHa.Open "SELECT * FROM `" & Tabl & "` LIMIT 1", MyCn, adOpenStatic, adLockReadOnly
       myFrag rHa, "SELECT * FROM `" & Tabl & "` LIMIT 1", adOpenStatic, MyCn
       If Not rHa.BOF Then
        db = rdb.Fields(0) ' 1.10.13: Führt mit ????k??a?k zum Fehler
        If InStrB(db, "") = 0 And InStrB(db, "") = 0 And InStrB(db, "?") = 0 Then
         eintragen = True
        End If
       End If
      End If ' not rha.bof
      rdb.Move 1
      If rdb.EOF And LenB(db) = 0 And obLeere Then eintragen = True
     Else
      eintragen = True
     End If ' lenb(tabl)<>0
     On Error GoTo fehler
     If eintragen Then
      Set acn(UBound(acn)) = New ADODB.Connection
      Dim VorLänge&
      VorLänge = Stri(0).length
      Stri(0).Cut (VorLänge - 1)
      Stri(0).AppVar Array(";DATABASE=", db, """")
      acn(UBound(acn)).ConnectionString = Stri(0).Value  ' Left$(MyCn.ConnectionString, Len(MyCn.ConnectionString) - 1) & ";DATABASE=" & db & """"
      Stri(0).Cut VorLänge
      acn(UBound(acn)).Open
      acSt(UBound(acSt)) = Stri(1)
      ReDim Preserve acSt(UBound(acSt) + 1)
      ReDim Preserve acn(UBound(acn) + 1)
     End If
     If LenB(Tabl) = 0 Then Exit Do
     If rdb.EOF Then Exit Do
     On Error Resume Next
    Loop
   End If ' err.number = 0
   On Error GoTo fehler
   End If
  Next Cpt
 Next runde
 If UBound(acn) <> 0 Then
  ReDim Preserve acn(UBound(acn) - 1)
  ReDim Preserve acSt(UBound(acSt) - 1)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & ", ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getAllDB/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getAllDB
