VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCltentry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Entry"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12150
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATA MANIPULATION"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   44
      Top             =   5040
      Width           =   6255
      Begin VB.CommandButton DELETE 
         BackColor       =   &H00FFFF80&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton UPDATE 
         BackColor       =   &H00FFFF80&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton SAVE 
         BackColor       =   &H00FFFF80&
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton ADD 
         BackColor       =   &H00FFFF80&
         Caption         =   "Add"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbType 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCltentry.frx":0000
      Left            =   7320
      List            =   "frmCltentry.frx":0002
      TabIndex        =   12
      Top             =   7560
      Width           =   4455
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   7920
      Width           =   4575
   End
   Begin VB.CommandButton SEARCH 
      BackColor       =   &H0080FF80&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtnutrition 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   8040
      Width           =   4455
   End
   Begin VB.TextBox txtfeet 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   8
      Top             =   1680
      Width           =   720
   End
   Begin VB.TextBox txtin 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.OptionButton optcm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Height in Cm"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton optfeet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Height in Feet"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cmbGender 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCltentry.frx":0004
      Left            =   2760
      List            =   "frmCltentry.frx":000E
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAVIGATION"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   28
      Top             =   6240
      Width           =   6255
      Begin VB.CommandButton LAST 
         BackColor       =   &H0080C0FF&
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton NEXT 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton PREV 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton FIRST 
         BackColor       =   &H0080C0FF&
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCltno 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   1440
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   7405569
      CurrentDate     =   43155
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtmobno 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtCltname 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8520
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDOV 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   7405569
      CurrentDate     =   43155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BMI Calculation"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7320
      TabIndex        =   29
      Top             =   2280
      Width           =   4455
      Begin VB.TextBox txtheight 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtweight 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblbmi 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/m^2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "In Kgs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "In Cms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Body Mass Index"
         BeginProperty Font 
            Name            =   "AR JULIAN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Height "
         BeginProperty Font 
            Name            =   "AR JULIAN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "AR JULIAN"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCLEAR 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   5640
      TabIndex        =   43
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblresult 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Client Name"
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
      Left            =   720
      TabIndex        =   40
      Top             =   7440
      Width           =   4215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Nutrition Suggested"
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
      Height          =   735
      Left            =   5640
      TabIndex        =   38
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Result :"
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
      Left            =   7440
      TabIndex        =   37
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   7320
      Picture         =   "frmCltentry.frx":0020
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Feet "
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Inches"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9480
      TabIndex        =   35
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Visit"
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
      TabIndex        =   24
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
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
      TabIndex        =   22
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      TabIndex        =   21
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Client Name"
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
      TabIndex        =   20
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Client No."
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
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client Entry"
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
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmCltentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Single

Private Sub ADD_Click()
sql = "select * from Client order by Client_No"
    Set RS2 = New ADODB.Recordset
    RS2.Open sql, CN, 1, 3
    If RS2.RecordCount = 0 Then
        txtCltno.Text = 1
    Else
        RS2.MoveLast
        txtCltno.Text = RS2.Fields("Client_No") + 1
    End If
    RS2.Close
ADD.Enabled = False
SAVE.Enabled = True
UPDATE.Enabled = False
DELETE.Enabled = False
SEARCH.Enabled = False
dtpDOB.Value = Date
dtpDOV.Value = Date
Call CLEAR
Call buttonEnabled
txtCltname.SetFocus
SAVE.Default = True

End Sub

Private Sub buttonEnabled()
txtCltno.Enabled = False
txtCltname.Enabled = True
dtpDOB.Enabled = True
cmbGender.Enabled = True
txtmobno.Enabled = True
txtadd.Enabled = True
dtpDOV.Enabled = True
txtweight.Enabled = True
txtnutrition.Enabled = True
optcm.Enabled = True
optfeet.Enabled = True
cmbType.Enabled = True
End Sub

Private Sub buttonDisabled()
txtCltno.Enabled = False
txtCltname.Enabled = False
dtpDOB.Enabled = False
cmbGender.Enabled = False
txtmobno.Enabled = False
txtadd.Enabled = False
dtpDOV.Enabled = False
txtweight.Enabled = False
txtnutrition.Enabled = False
optcm.Enabled = False
optfeet.Enabled = False
cmbType.Enabled = False
End Sub

Private Sub cmbtype_Click()
sql = "select * from Unbalanced where Type='" & cmbType.Text & "'"
Set RS1 = New ADODB.Recordset
RS1.Open sql, CN, 1, 3
 txtnutrition.Text = RS1.Fields(1)
End Sub

Private Sub fill()
   cmbType.CLEAR

Do While Not RS1.EOF
   cmbType.AddItem RS1!Type
    RS1.MoveNext
Loop
End Sub

Private Sub cmdBack_Click()
    Unload Me
    MDIForm1.mnudef = True
    MDIForm1.mnusearch = True
    MDIForm1.mnuview = True
    MDIForm1.mnuexit = True
End Sub

Private Sub DELETE_Click()
Dim YN As Integer
If txtCltname.Text = "" Or txtweight.Text = "" Then
    MsgBox "Record Not Selected", vbCritical, "Diet Hub"
    Else
YN = MsgBox("ARE YOU SURE", vbYesNo + vbQuestion, "Diet Hub")
If YN = vbYes Then
RS2.DELETE
MsgBox "Record Deleted Successfully", vbInformation, "Diet Hub"
RS2.MoveNext
If RS2.EOF Then
   RS2.MoveLast
End If
Call showdata
End If
End If
    
End Sub

Private Sub FIRST_Click()
sql = "select * from Client"
        Set RS2 = New ADODB.Recordset
        RS2.Open sql, CN, 1, 3
RS2.MoveFirst
UPDATE.Enabled = True
DELETE.Enabled = True
SAVE.Enabled = False
ADD.Enabled = False
Call showdata
Call buttonEnabled
End Sub

Private Sub LAST_Click()
RS2.MoveLast
    Call showdata
End Sub

Private Sub NEXT_Click()
RS2.MoveNext
If Not RS2.EOF Then
    Call showdata
Else
    MsgBox "Last Record", vbInformation, "Diet Hub"
End If
End Sub

Private Sub optcm_Click()
    If optcm.Value = True Then
        txtheight.Enabled = True
        txtin.Enabled = False
        txtfeet.Enabled = False
        txtheight.SetFocus
        txtfeet.Text = ""
        txtin.Text = ""
        txtheight = ""
    End If
End Sub

Private Sub optfeet_Click()
    If optfeet.Value = True Then
        txtheight.Enabled = False
        txtin.Enabled = True
        txtfeet.Enabled = True
        txtfeet.SetFocus
        txtheight.Text = ""
    End If
End Sub

Private Sub PREV_Click()
RS2.MovePrevious
    If Not RS2.BOF Then
        Call showdata
    Else
        MsgBox "First Record", vbInformation, "Diet Hub"
    End If
End Sub

Private Sub SEARCH_Click()
sql = "select * from Client where Client_Name='" + txtSearch.Text + "'"
Set RS2 = New ADODB.Recordset
RS2.Open sql, CN, 1, 3
    If RS2.EOF Then
        MsgBox "Client Not Found", vbCritical, "Diet Hub"
    Else
        Call showdata
        Call buttonEnabled
    End If
UPDATE.Enabled = True
DELETE.Enabled = True
End Sub

Private Sub txtCltname_Keypress(KeyAscii As Integer)
    If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 32 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Alphabets", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtin_Change()
    txtheight.Text = (Val(txtfeet.Text) * 12 + Val(txtin.Text)) * 2.54
End Sub

Private Sub cmdCLEAR_Click()
    txtCltno = ""
    Call CLEAR
    ADD.Enabled = True
    SEARCH.Enabled = True
    UPDATE.Enabled = False
    DELETE.Enabled = False
    SAVE.Enabled = False
    Call buttonDisabled
End Sub

Private Sub txtmobno_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Numbers", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtfeet_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Numbers", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtin_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Numbers", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtSearch_Change()
    SEARCH.Default = True
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or KeyAscii = 32 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Alphabets", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtweight_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        MsgBox "Enter Only Numbers", vbExclamation, "Diet Hub"
    End If
End Sub

Private Sub txtweight_Change()
If txtheight.Text = "" Then

Else
    a = Val(txtheight.Text) / 100
    lblbmi.Caption = Val(txtweight.Text) / (a * a)
    If Val(lblbmi.Caption) <= 18.5 Then
    lblresult.Caption = "Under_Weight"
    ElseIf Val(lblbmi.Caption) >= 18.5 And Val(lblbmi.Caption) <= 24.9 Then
    
    lblresult.Caption = "Normal_Weight"
    ElseIf Val(lblbmi.Caption) >= 25 And Val(lblbmi.Caption) <= 29.9 Then
    lblresult.Caption = "Over_Weight"
    ElseIf Val(lblbmi.Caption) > 30 Then
    lblresult.Caption = "Obesity"
    
    End If
End If
End Sub

Private Sub UPDATE_Click()
If txtCltname.Text = "" Or txtweight.Text = "" Then
    MsgBox "Record Not Selected", vbCritical, "Diet Hub"
Else
        sql = "select * from Client where Client_No=" & txtCltno.Text
        Set RS2 = New ADODB.Recordset
        RS2.Open sql, CN, 1, 3
         RS2.Fields("Client_Name") = txtCltname.Text
         RS2.Fields("DOB") = dtpDOB.Value
         RS2.Fields("Gender") = cmbGender.Text
         RS2.Fields("Mobile_No") = txtmobno.Text
         RS2.Fields("Address") = txtadd.Text
         RS2.Fields("DOV") = dtpDOV.Value
         RS2.Fields("Height") = txtheight.Text
         RS2.Fields("Weight") = txtweight.Text
         RS2.Fields("BMI") = lblbmi.Caption
         RS2.Fields("Result") = lblresult.Caption
         RS2.Fields("Type") = cmbType.Text
         RS2.Fields("Nutrition") = txtnutrition.Text
        RS2.UPDATE
        RS2.Close
        MsgBox "Record Update Successfully", vbInformation, "Diet Hub"
        Unload Me
ADD.Enabled = True
SEARCH.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Call connect
    Call buttonDisabled
    txtCltno = ""
    Call CLEAR
    Call fill
    dtpDOB.Value = Date
    dtpDOV.Value = Date
    ADD.Default = True
    cmdBack.Default = False
End Sub

Private Sub showdata()
If Not IsNull(RS2.Fields(0)) Then
    txtCltno = RS2.Fields(0)
End If
If Not IsNull(RS2.Fields(1)) Then
    txtCltname = RS2.Fields(1)
End If
If Not IsNull(RS2.Fields(2)) Then
    dtpDOB = RS2.Fields(2)
End If
If Not IsNull(RS2.Fields(3)) Then
    cmbGender = RS2.Fields(3)
End If
If Not IsNull(RS2.Fields(4)) Then
    txtmobno = RS2.Fields(4)
End If
If Not IsNull(RS2.Fields(5)) Then
    txtadd = RS2.Fields(5)
End If
If Not IsNull(RS2.Fields(6)) Then
    dtpDOV = RS2.Fields(6)
End If
If Not IsNull(RS2.Fields(7)) Then
    txtheight = RS2.Fields(7)
End If
If Not IsNull(RS2.Fields(8)) Then
    txtweight = RS2.Fields(8)
End If
If Not IsNull(RS2.Fields(9)) Then
    lblbmi = RS2.Fields(9)
End If
If Not IsNull(RS2.Fields(10)) Then
    lblresult = RS2.Fields(10)
End If
If Not IsNull(RS2.Fields(11)) Then
    cmbType = RS2.Fields(11)
End If
If Not IsNull(RS2.Fields(12)) Then
    txtnutrition = RS2.Fields(12)
End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    CN.Close
End Sub

Private Sub CLEAR()
txtCltname = ""
cmbGender = ""
txtmobno = ""
txtadd = ""
txtfeet = ""
txtin = ""
txtheight = ""
txtweight = ""
lblbmi = ""
lblresult = ""
cmbType = ""
txtnutrition = ""
txtSearch = ""
End Sub

Private Sub SAVE_Click()
sql = "select * from Client where Client_No=" & txtCltno.Text
Set RS2 = New ADODB.Recordset
RS2.Open sql, CN, 1, 3
If txtCltname.Text = "" Or txtweight.Text = "" Then
    MsgBox "Fill All the Entries", vbInformation, "Diet Hub"
Else
         RS2.AddNew
         RS2.Fields("Client_No") = txtCltno.Text
         RS2.Fields("Client_Name") = txtCltname.Text
         RS2.Fields("DOB") = dtpDOB.Value
         RS2.Fields("Gender") = cmbGender.Text
         RS2.Fields("Mobile_No") = txtmobno.Text
         RS2.Fields("Address") = txtadd.Text
         RS2.Fields("DOV") = dtpDOV.Value
         RS2.Fields("Height") = txtheight.Text
         RS2.Fields("Weight") = txtweight.Text
         RS2.Fields("BMI") = lblbmi.Caption
         RS2.Fields("Result") = lblresult.Caption
         RS2.Fields("Type") = cmbType.Text
         RS2.Fields("Nutrition") = txtnutrition.Text
        RS2.UPDATE
        RS2.Close
        MsgBox "Entry Successful", vbInformation, "Diet Hub"
        Unload Me
SAVE.Enabled = False
ADD.Enabled = True
UPDATE.Enabled = True
SEARCH.Enabled = True
End If
ADD.Default = True
End Sub
