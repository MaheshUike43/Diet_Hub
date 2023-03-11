VERSION 5.00
Begin VB.Form frmsearchdef 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Deficiency"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10980
   Begin VB.ComboBox cmbtype 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   420
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   8055
   End
   Begin VB.CommandButton BACK 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   2415
   End
   Begin VB.TextBox txtdeff 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   8055
   End
   Begin VB.TextBox txtdoses 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5280
      Width           =   8055
   End
   Begin VB.TextBox txtnutri 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Deficiency :"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Doses :"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search Deficiency"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nutrients :"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmsearchdef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbtype_Click()
S = "select * from Unbalanced where Type='" & cmbType.Text & "'"
Set RS1 = New ADODB.Recordset
RS1.Open S, CN, 1, 3
 txtnutri.Text = RS1.Fields(1)
 txtdeff.Text = RS1.Fields(2)
 txtdoses.Text = RS1.Fields(3)
End Sub

Private Sub BACK_Click()
    Unload Me
    MDIForm1.mnuCltdata = True
    MDIForm1.mnusdet = True
    MDIForm1.mnuentry = True
    MDIForm1.mnuview = True
    MDIForm1.mnuexit = True
End Sub

Private Sub Form_Load()
    Call connect
    Call fill
    
End Sub
Private Sub fill()
   cmbType.CLEAR

Do While Not RS1.EOF
   cmbType.AddItem RS1!Type
    RS1.MoveNext
Loop
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    CN.Close
End Sub

