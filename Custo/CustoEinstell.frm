VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Const Cn$ = "Driver={Microsoft Access Driver (*.mdb)};Dbq=u:\custobase.mdb;Jet OLEDB:Database Password='CDS'"
Const Cn$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=u:\custobase.mdb;Jet OLEDB:Database Password=CDS"
Const LVobMySQL% = 0
Private Sub Command1_Click()
 Dim acn As New ADODB.connection, rs As New Recordset, i&
 acn.Open Cn
 acn.Execute ("update systemdaten set [WI-Datum] = #04/22/2018 23:59:59#, [WI-Level]=0, [WI-LevelIncementDate] = #04/22/2018 23:59:59# where [WI-Level] > 0")
 rs.Open "systemdaten", Cn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  For i = 0 To rs.Fields.Count - 1
   Debug.Print i, rs.Fields(i).Name, rs.Fields(i)
  Next i
  Debug.Print " "
  rs.Move 1
 Loop
End Sub

