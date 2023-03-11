VERSION 5.00
Begin VB.Form frmdef 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutrition Deficiency"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   15915
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
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6000
      Width           =   1215
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
      Left            =   8880
      TabIndex        =   21
      Top             =   5400
      Width           =   4575
   End
   Begin VB.CommandButton BACK 
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
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
      Height          =   855
      Left            =   7800
      TabIndex        =   2
      Top             =   1320
      Width           =   7815
   End
   Begin VB.TextBox txttype 
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
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Frame Frame3 
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
      Height          =   1200
      Left            =   1920
      TabIndex        =   15
      Top             =   4560
      Width           =   5865
      Begin VB.CommandButton ADD 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "To Add Record"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton SAVE 
         BackColor       =   &H0080FF80&
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "to save  record"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton UPDATE 
         BackColor       =   &H0080FF80&
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "to edit record"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton DELETE 
         BackColor       =   &H0080FF80&
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "to delete record"
         Top             =   480
         Width           =   1095
      End
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
      Height          =   1335
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Frame FRAME2 
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
      Height          =   1215
      Left            =   1920
      TabIndex        =   10
      Top             =   5880
      Width           =   5835
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
         Index           =   0
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "To View First Record"
         Top             =   480
         Width           =   1140
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
         Index           =   1
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "View Previous Record"
         Top             =   480
         Width           =   1155
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
         Index           =   2
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "View Next Record"
         Top             =   480
         Width           =   1245
      End
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
         Index           =   3
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "View Last Record"
         Top             =   480
         Width           =   1155
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancel Operation"
      Top             =   6000
      Width           =   1335
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
      Height          =   1335
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Type"
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
      Left            =   9000
      TabIndex        =   23
      Top             =   4920
      Width           =   4215
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
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nutrition Deficiency"
      BeginProperty Font 
         Name            =   "Magneto"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmdef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ADD_Click()
ADD.Enabled = False
SAVE.Enabled = True
UPDATE.Enabled = False
DELETE.Enabled = False
Call CLEAR
RS1.MoveLast
Call buttonEnabled
txttype.SetFocus
SAVE.Default = True
SEARCH.Enabled = False
End Sub

Private Sub buttonEnabled()
    txttype.Enabled = True
    txtnutri.Enabled = True
    txtdeff.Enabled = True
    txtdoses.Enabled = True
End Sub

Private Sub buttonDisabled()
    txttype.Enabled = False
    txtnutri.Enabled = False
    txtdeff.Enabled = False
    txtdoses.Enabled = False
End Sub

Private Sub BACK_Click()
    Unload Me
    MDIForm1.mnuClte = True
    MDIForm1.mnusearch = True
    MDIForm1.mnuview = True
    MDIForm1.mnuexit = True
End Sub

Private Sub DELETE_Click()
Dim YN As Integer
If txttype.Text = "" Or txtdoses.Text = "" Then
    MsgBox "Record Not Selected", vbCritical, "Diet Hub"
    Else
YN = MsgBox("ARE YOU SURE", vbYesNo + vbQuestion, "Diet Hub")
If YN = vbYes Then
RS1.DELETE
MsgBox "Record Deleted Successfully", vbInformation, "Diet Hub"
RS1.MoveNext
If RS1.EOF Then
   RS1.MoveLast
End If
Call showdata
End If
End If
UPDATE.Enabled = False
End Sub

Private Sub FIRST_Click(Index As Integer)
sql = "select * from Unbalanced"
        Set RS1 = New ADODB.Recordset
        RS1.Open sql, CN, 1, 3
RS1.MoveFirst
UPDATE.Enabled = True
DELETE.Enabled = True
SAVE.Enabled = False
ADD.Enabled = False
Call showdata
Call buttonEnabled
End Sub

Private Sub LAST_Click(Index As Integer)
RS1.MoveLast
    Call showdata
End Sub

Private Sub NEXT_Click(Index As Integer)
RS1.MoveNext
If Not RS1.EOF Then
    Call showdata
Else
    MsgBox "Last Record", vbInformation, "Diet Hub"
End If
End Sub

Private Sub PREV_Click(Index As Integer)
RS1.MovePrevious
    If Not RS1.BOF Then
        Call showdata
    Else
        MsgBox "First Record", vbInformation, "Diet Hub"
    End If
End Sub

Private Sub cmdCLEAR_Click()
Call CLEAR
ADD.Enabled = True
UPDATE.Enabled = False
DELETE.Enabled = False
SAVE.Enabled = False
SEARCH.Enabled = True
Call buttonDisabled
End Sub


Private Sub SEARCH_Click()
    sql = "select * from Unbalanced where Type='" + txtSearch.Text + "'"
    Set RS1 = New ADODB.Recordset
    RS1.Open sql, CN, 1, 3
    If RS1.EOF Then
        MsgBox "Type Not Found", vbCritical, "Diet Hub"
    Else
        Call showdata
        Call buttonEnabled
    End If
    UPDATE.Enabled = True
    DELETE.Enabled = True
End Sub

Private Sub txtSearch_Change()
    SEARCH.Default = True
End Sub

Private Sub UPDATE_Click()
If txttype.Text = "" Then
    MsgBox "Record Not Selected", vbCritical, "Diet Hub"
Else
sql = "select * from Unbalanced where Type='" + txttype.Text + "'"
Set RS1 = New ADODB.Recordset
RS1.Open sql, CN, 1, 3
         RS1.Fields("Type") = txttype.Text
         RS1.Fields("Nutrients") = txtnutri.Text
         RS1.Fields("Deficiency") = txtdeff.Text
         RS1.Fields("Doses") = txtdoses.Text
        RS1.UPDATE
        RS1.Close
        MsgBox "Record Update Successfully", vbInformation, "Diet Hub"
        Unload Me
ADD.Enabled = True
SEARCH.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Call connect
    Call buttonDisabled
    Call CLEAR
    ADD.Default = True
End Sub

Private Sub showdata()
If Not IsNull(RS1.Fields(0)) Then
    txttype = RS1.Fields(0)
End If
If Not IsNull(RS1.Fields(1)) Then
    txtnutri = RS1.Fields(1)
End If
If Not IsNull(RS1.Fields(2)) Then
    txtdeff = RS1.Fields(2)
End If
If Not IsNull(RS1.Fields(3)) Then
    txtdoses = RS1.Fields(3)
End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    CN.Close
End Sub

Private Sub CLEAR()
txttype = ""
txtnutri = ""
txtdeff = ""
txtdoses = ""
txtSearch = ""
End Sub

Private Sub SAVE_Click()
sql = "select * from Unbalanced"
Set RS1 = New ADODB.Recordset
RS1.Open sql, CN, 1, 3
If txttype.Text = "" Or txtnutri.Text = "" Or txtdeff.Text = "" Or txtdoses.Text = "" Then
    MsgBox "Fill All the Entries", vbInformation, "Diet Hub"
    Else
         RS1.AddNew
         RS1.Fields("Type") = txttype.Text
         RS1.Fields("Nutrients") = txtnutri.Text
         RS1.Fields("Deficiency") = txtdeff.Text
         RS1.Fields("Doses") = txtdoses.Text
        RS1.UPDATE
        RS1.Close
        MsgBox "Entry Successful", vbInformation, "Diet Hub"
        Unload Me
SAVE.Enabled = False
ADD.Enabled = True
UPDATE.Enabled = True
SEARCH.Enabled = True
End If
End Sub
