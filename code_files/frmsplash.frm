VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIET HUB"
   ClientHeight    =   10830
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   19125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   19125
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   4800
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Miss. Jayshri S. Aswale"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   4
      Top             =   9240
      Width           =   9015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mr. Swapnil V. Fale"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   2
      Top             =   8040
      Width           =   8175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By :"
      BeginProperty Font 
         Name            =   "Cooper Std Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Diet Hub"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1575
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mr. Mahesh B. Uike"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   3
      Top             =   6840
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   16005
      Left            =   -1680
      Picture         =   "frmsplash.frx":0000
      Top             =   -1440
      Width           =   24000
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Unload Me
    frmLogin.Show
End Sub
