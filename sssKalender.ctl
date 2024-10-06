VERSION 5.00
Begin VB.UserControl sssKalender 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ScaleHeight     =   3810
   ScaleWidth      =   2775
   Begin VB.ListBox lstYear 
      Height          =   2400
      Left            =   1200
      TabIndex        =   60
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdJvor 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "ein Jahr nach vorne"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdMvor 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "einen Monat nach vorne"
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdMzur 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "einen Monat zurück"
      Top             =   600
      Width           =   255
   End
   Begin VB.ListBox lstMonth 
      Height          =   2400
      ItemData        =   "sssKalender.ctx":0000
      Left            =   840
      List            =   "sssKalender.ctx":002B
      TabIndex        =   58
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdJzur 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "ein Jahr zurück"
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   46
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   1
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   45
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   2
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   44
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   3
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   43
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   4
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   42
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   5
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   41
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   6
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   40
      Top             =   1275
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   7
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   39
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   8
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   38
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   9
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   37
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   10
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   36
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   11
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   35
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   12
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   34
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   13
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   33
      Top             =   1635
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   14
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   32
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   15
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   31
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   16
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   30
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   17
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   29
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   18
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   28
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H80000014&
      Caption         =   "30"
      Height          =   375
      Index           =   19
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   27
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   20
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   26
      Top             =   1995
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   21
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   25
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   22
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   24
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   23
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   23
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   24
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   22
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   25
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   21
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   26
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   20
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   27
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   19
      Top             =   2355
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   28
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   18
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   29
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   17
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   30
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   16
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   31
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   15
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   32
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   14
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   33
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   13
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   34
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   2715
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   35
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   36
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   37
      Left            =   840
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   38
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   39
      Left            =   1560
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   40
      Left            =   1920
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   3075
      Width           =   375
   End
   Begin VB.OptionButton optDay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "30"
      Height          =   375
      Index           =   41
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   3075
      Width           =   375
   End
   Begin VB.CommandButton cmdDrop 
      Height          =   320
      Left            =   2280
      Picture         =   "sssKalender.ctx":008E
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Kalender aufklappen / schließen"
      Top             =   0
      Width           =   320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   45
      TabIndex        =   47
      Top             =   1215
      Width           =   2700
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   3300
      Left            =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "1998"
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   1630
      TabIndex        =   59
      ToolTipText     =   "Jahres-Anzeige und -Direktwahl"
      Top             =   645
      Width           =   375
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Oktober"
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   720
      TabIndex        =   57
      ToolTipText     =   "Monats-Anzeige u. -Direktwahl"
      Top             =   645
      Width           =   855
   End
   Begin VB.Shape Shape1 
      DrawMode        =   16  'Stift mischen
      Height          =   3600
      Index           =   0
      Left            =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblDatAnz 
      Alignment       =   2  'Zentriert
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      Height          =   300
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblToday 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "heute ist der"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   55
      Top             =   3560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "MO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "DI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   53
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "MI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   52
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "DO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   51
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "FR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   50
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "SA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   49
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "SO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   48
      Top             =   915
      Width           =   375
   End
End
Attribute VB_Name = "sssKalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'sssKalender.ocx
'Datumssteuerelement mit aufklappbarem Kalender
'
'Original by Peter Preintner (Preintner.P@t-online.de)
'
'Dieses OCX ist Freeware, solange es privat benutzt wird.
'
'Wenn jemand Ideen hat, wie man dieses Steuerelement noch verbessern
'kann, dann sendet den neuen Code bitte an mich zurück.
'
'Viel Spaß!  :-)


Option Explicit
Dim aktDate As Date
Dim aktzp#
Dim aktDay As Integer
Dim aktMonth As Integer
Dim aktYear As Integer
Dim sTag As String
Dim sMonat As String
Dim sJahr As String
Dim zJahr As Integer
Dim zTag As Integer
Dim zMonat As Integer
Dim i As Integer

'Kinstanten für Eigenschaftswerte:
Const m_def_Enabled = True
Const m_def_Value = "01.01.1998"
Const m_def_Cancel = 0
Const HoeheZU = 315
Const HoeheAUF = 3810

'Eigenschaft-Variablen:
Dim m_Value As Date
Dim m_Enabled As Boolean
Dim m_Cancel As Boolean

'Ereignisdeklarationen:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Tritt auf, wenn ein Formular zum ersten Mal angezeigt wird oder wenn sich die Größe eines Objekts ändert."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Tritt auf, wenn sich die Visible-Eigenschaft des Steuerelements in True ändert."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus besitzt."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus besitzt."
Public Event Change()
Public Event Closed()


Private Sub cmdDrop_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdDrop_Click()
On Error GoTo fehler
For i = 0 To lstYear.ListCount - 1
    If lstYear.List(i) = CStr(Format(m_Value, "yyyy")) Then
      lstYear.ListIndex = i
      Exit For
    End If
Next i
If UserControl.Height > 500 Then
  UserControl.Height = 315
  RaiseEvent Closed
  Exit Sub
End If
UserControl.Height = 3810
Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDropClick/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdDrop_Click

Private Sub cmdDropUp_Click(Index%)
 On Error GoTo fehler
 UserControl.Height = 495
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in cmdDropUp_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdDropUp_Click

Private Sub lstMonth_Click()
On Error GoTo fehler0
  aktMonth = lstMonth.ItemData(lstMonth.ListIndex)
  lstMonth.Visible = False
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  aktzp = aktDate
Exit Sub
fehler0:
Select Case Err.Number
  'wenn zB. 30. im Februar
  Case 13
    'Monat mit 30 Tagen
    If aktDay > 30 Then aktDay = 30: Resume
    'Monat mit 29 Tagen (Schaltjahr)
    If aktDay > 29 Then aktDay = 29: Resume
    'Monat mit 28 Tagen
    If aktDay > 28 Then aktDay = 28: Resume
End Select
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in lstMonth_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' lstMonth_Click

Private Sub lstYear_Click()
 On Error GoTo fehler
  aktYear = lstYear.List(lstYear.ListIndex)
  lstYear.Visible = False
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  aktzp = aktDate
  lblYear.Caption = Format(aktDate, "yyyy")
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in lstYear_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' lstYear_Click

Private Sub lblMonth_Click()
  On Error GoTo fehler
  lstMonth.Visible = True
  lstYear.Visible = False
 Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in lblMonth_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' lblMonth_Click

Private Sub lblYear_Click()
  On Error GoTo fehler
  lstYear.Visible = True
  lstMonth.Visible = False
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in lblYear_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' lblYear_Click

Private Sub UserControl_Initialize()
 On Error GoTo fehler
  DateZerleger Now
  lblToday.Caption = "Heute ist der " & Format(zTag, "00") & ". " & sMonat & " " & sJahr
  aktDate = Now
  aktzp = aktDate
  aktYear = Format(aktDate, "yyyy")
  Monatswahl aktDate
  TagesAuswahl aktDate
  DateZerleger Now
  DatumsEintrag
  For i = 1960 To 2100
    lstYear.AddItem i
  Next i
  UserControl.Height = 315
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in UserControl_Initialize/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' UserControl_Initialize

Private Sub ValueEintrag()
 On Error GoTo fehler
  Monatswahl m_Value
  TagesAuswahl m_Value
  DateZerleger m_Value
  DatumsEintrag
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ValueEintrag/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

'
Private Sub Monatswahl(Datum As Date)
  Dim Starttag As String
  Dim startIndex As Integer
  Dim StartDate As Date
  Dim Counter As Integer
  Dim VorDate As Date
  Dim Färber As Boolean
  
  On Error GoTo fehler
  Färber = False
  Starttag = ErsterMonatsTag(Datum)
  StartDate = CDate("01." & CStr(Month(Datum)) _
      & "." & CStr(Format(Datum, "yyyy")))
  'alle auf weiß setzen
  For i = 0 To optDay.Count - 1
    optDay(i).BackColor = &HFFFFFF
  Next i
  Select Case Starttag
    Case "Montag"
      startIndex = 7
    Case "Dienstag"
      startIndex = 8
    Case "Mittwoch"
      startIndex = 9
    Case "Donnerstag"
      startIndex = 3
    Case "Freitag"
      startIndex = 4
    Case "Samstag"
      startIndex = 5
    Case "Sonntag"
      startIndex = 6
  End Select
  'Beschriften der Tagesfelder
  DateZerleger StartDate
  VorDate = CDate("26." & zMonat & "." & zJahr)
  Counter = 0
  While Day(VorDate) <> 1
    Counter = Day(VorDate)
    VorDate = DateAdd("d", 1, VorDate)
  Wend

  For i = startIndex - 1 To 0 Step -1
    optDay(i).Caption = Counter
    'grauen
    optDay(i).BackColor = &H8000000F
    optDay(i).Tag = 0
    Counter = Counter - 1
  Next i
  Counter = 0
  For i = startIndex To 41
    optDay(i).Caption = Day(DateAdd("d", Counter, StartDate))
    optDay(i).Tag = 1
    'grauen
    If Day(DateAdd("d", Counter, StartDate)) = 1 And Counter > 5 Then
      Färber = True
    End If
    If Färber = True Then
      optDay(i).BackColor = &H8000000F
      optDay(i).Tag = 0
    End If
    Counter = Counter + 1
  Next i
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Monatswahl/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Monatswahl

Private Sub TagesAuswahl(Datum As Date)
  On Error GoTo fehler
  For i = 0 To optDay.Count - 1
   If optDay(i).Tag = 1 Then
    If optDay(i).Caption = CStr(Day(Datum)) Then
        optDay(i).BackColor = &H800000
        optDay(i).Value = True
        GoTo Nextes
    End If
    optDay(i).BackColor = &HFFFFFF
Nextes:
   End If
  Next i
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in TagesAuswahl/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' TagesAuswahl

Private Function DateZerleger(Datum As Date)
  zTag = Day(Datum)
  zMonat = Month(Datum)
  zJahr = Year(Datum)
  sTag = CStr(Format(Datum, "dddd"))
  sMonat = CStr(Format(Datum, "mmmm"))
  sJahr = CStr(Format(Datum, "yyyy"))
  aktDay = Day(Datum)
  aktMonth = Month(Datum)
  aktYear = Year(Datum)
  aktzp = Datum
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DateZerleger/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DateZerleger

Private Function ErsterMonatsTag(Datum As Date)
  Dim StartDate As Date
  On Error GoTo fehler
  StartDate = CDate("01." & CStr(Month(Datum)) _
      & "." & CStr(Format(Datum, "yyyy")))
  ErsterMonatsTag = CStr(Format(StartDate, "dddd"))
  Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in ErsterMonatsTag/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ErsterMonatsTag

Private Sub DatumsEintrag()
 On Error GoTo fehler
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  lblDatAnz.Caption = Format(aktDate, "ddd") & "  " & Format(aktzp, "dd.mm.yyyy hh:mm:ss")
  lblMonth.Caption = Format(aktDate, "mmmm")
  lblMonth.Refresh
  m_Value = CDate(aktDay & "." & aktMonth & "." & aktYear)
  lstMonth.ListIndex = aktMonth - 1
  For i = 0 To lstYear.ListCount - 1
    If lstYear.List(i) = CStr(Format(m_Value, "yyyy")) Then
      lstYear.ListIndex = i
      Exit For
    End If
  Next i
  RaiseEvent Change
  For i = 0 To optDay.Count - 1
   If optDay(i).Tag = 1 Then
    If optDay(i).Caption = CStr(Day(aktDate)) Then
        optDay(i).BackColor = &H800000
        GoTo Nextes
    End If
    optDay(i).BackColor = &HFFFFFF
Nextes:
   End If
  Next i
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in DatumsEintrag/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DatumsEintrag

Private Sub optDay_Click(Index As Integer)
  'wenn gewählter Tag im Monatsbereich
  If optDay(Index).Tag = 1 Then
    aktDay = CInt(optDay(Index).Caption)
    aktzp = CDate(aktDay & "." & aktMonth & "." & aktYear)
    DatumsEintrag
  Else
    If Index < 20 Then cmdMzur_Click
    If Index > 20 Then cmdMvor_Click
  End If
  Exit Sub
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in optDay_Click/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' optDay_Click

Private Sub cmdMvor_Click()
On Error GoTo fehler
  'Monat weiter
  If aktMonth = 12 Then aktMonth = 0
  aktMonth = aktMonth + 1
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  aktzp = aktDate
  'Datumseintrag
Exit Sub
fehler:
Select Case Err.Number
  'wenn zB. 30. im Februar
  Case 13
    If aktDay > 28 Then aktDay = 28
    Resume
End Select
End Sub

Private Sub cmdMzur_Click()
On Error GoTo fehler
  'Monat zurück
  If aktMonth = 1 Then aktMonth = 13
  aktMonth = aktMonth - 1
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  aktzp = aktDate
  'Datumseintrag
Exit Sub
fehler:
Select Case Err.Number
  'wenn zB. 30. im Februar
  Case 13
    'Monat mit 30 Tagen
    If aktDay > 30 Then aktDay = 30: Resume
    'Monat mit 29 Tagen (Schaltjahr)
    If aktDay > 29 Then aktDay = 29: Resume
    'Monat mit 30 Tagen
    If aktDay > 28 Then aktDay = 28: Resume
End Select
End Sub

Private Sub cmdJvor_Click()
  'Jahr vor
  aktYear = aktYear + 1
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  'Datumseintrag
End Sub

Private Sub cmdJzur_Click()
  'Jahr vor
  aktYear = aktYear - 1
  aktDate = CDate(aktDay & "." & aktMonth & "." & aktYear)
  Monatswahl aktDate
  DoEvents
  TagesAuswahl aktDate
  aktzp = aktDate
  'Datumseintrag
End Sub


Private Sub UserControl_Resize()
  RaiseEvent Resize
  UserControl.Width = 2780
'  If Not UserControl.Height = HoeheZU Or UserControl.Height = HoeheAUF Then
'    UserControl.Height = HoeheZU
'  End If
End Sub

Private Sub UserControl_Show()
  RaiseEvent Show
End Sub

Public Property Get Value() As Date
Attribute Value.VB_Description = "Gibt den Wert eines Objekts zurück oder legt diesen fest."
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Date)
  m_Value = New_Value
  PropertyChanged "Value"
  ValueEintrag
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Public Property Get Cancel() As Boolean
Attribute Cancel.VB_Description = "Bestimmt, ob eine Befehlsschaltfläche die Schaltfläche ""Abbrechen"" in einem Formular ist."
  Cancel = m_Cancel
End Property

Public Property Let Cancel(ByVal New_Cancel As Boolean)
  m_Cancel = New_Cancel
  PropertyChanged "Cancel"
End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Gibt die Anzahl der Einheiten für die vertikale Messung des Inneren eines Objekts zurück oder legt diese fest."
  ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
  UserControl.ScaleHeight() = New_ScaleHeight
  PropertyChanged "ScaleHeight"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Gibt die horizontalen Koordinaten für die linken Ränder eines Objekts an oder legt diese fest."
  ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
  UserControl.ScaleLeft() = New_ScaleLeft
  PropertyChanged "ScaleLeft"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Gibt einen Wert zurück, der die Maßeinheiten für Objektkoordinaten bestimmt, wenn Grafikmethoden verwendet oder Steuerelemente positioniert werden, oder legt diesen fest."
  ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
  UserControl.ScaleMode() = New_ScaleMode
  PropertyChanged "ScaleMode"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Gibt die vertikalen Koordinaten für die oberen Ränder eines Objekts zurück oder legt diese fest."
  ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
  UserControl.ScaleTop() = New_ScaleTop
  PropertyChanged "ScaleTop"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Gibt die Anzahl der Einheiten für die horizontalen Abmessungen des Inneren eines Objekts zurück oder legt diese fest."
  ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
  UserControl.ScaleWidth() = New_ScaleWidth
  PropertyChanged "ScaleWidth"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function ScaleX(Width As Single, Optional FromScale As Variant, Optional ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Konvertiert den Wert für die Breite eines Formulars, Bildfeldes oder einer Druckerausgabe von einer Maßeinheit in eine andere."
  ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function ScaleY(Height As Single, Optional FromScale As Variant, Optional ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Konvertiert den Wert für die Höhe eines Formulars, Bildfeldes oder einer Druckerausgabe von einer Maßeinheit in eine andere."
  ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function


'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
  m_Enabled = m_def_Enabled
  m_Value = m_def_Value
  m_Cancel = m_def_Cancel
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  m_Cancel = PropBag.ReadProperty("Cancel", m_def_Cancel)
  On Error Resume Next
  UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3525)
  UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
  UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
  UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
  UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 3300)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("Cancel", m_Cancel, m_def_Cancel)
  Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3525)
  Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
  Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
  Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
  Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 3300)
End Sub


Public Function Zuklappen()
  UserControl.Height = HoeheZU
End Function

Public Function Aufklappen()
  UserControl.Height = HoeheAUF
End Function

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
  m_Enabled = vNewValue
  PropertyChanged "Enabled"
  Select Case m_Enabled
   Case 0
    lblDatAnz.Enabled = False
    cmdDrop.Enabled = False
   Case 1
    lblDatAnz.Enabled = True
    cmdDrop.Enabled = True
  End Select
End Property
