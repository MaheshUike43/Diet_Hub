VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Diet Hub"
   ClientHeight    =   7350
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   7335
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuentry 
      Caption         =   "&Entry"
      Begin VB.Menu mnuClte 
         Caption         =   "Client"
      End
      Begin VB.Menu mnudef 
         Caption         =   "Deficiency"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuCltdata 
         Caption         =   "Client Data"
      End
      Begin VB.Menu mnusd 
         Caption         =   "Deficiency"
      End
      Begin VB.Menu mnusdet 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnudefiency 
         Caption         =   "Deficiency"
      End
      Begin VB.Menu mnudet 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClte_Click()
    frmCltentry.Show
    mnudef.Enabled = False
    MDIForm1.mnusearch = False
    MDIForm1.mnuview = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnudef_Click()
    frmdef.Show
    mnuClte.Enabled = False
    MDIForm1.mnusearch = False
    MDIForm1.mnuview = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnuCltdata_Click()
    frmSviewClt.Show
    mnusd.Enabled = False
    mnusdet.Enabled = False
    MDIForm1.mnuentry = False
    MDIForm1.mnuview = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnusd_Click()
    frmsearchdef.Show
    mnuCltdata.Enabled = False
    mnusdet.Enabled = False
    MDIForm1.mnuentry = False
    MDIForm1.mnuview = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnusdet_Click()
    frmsearchdet.Show
    mnuCltdata.Enabled = False
    mnusd.Enabled = False
    MDIForm1.mnuentry = False
    MDIForm1.mnuview = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnudefiency_Click()
    frmviewdef.Show
    mnudet.Enabled = False
    MDIForm1.mnuentry = False
    MDIForm1.mnusearch = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnudet_Click()
    frmviewdet.Show
    mnudefiency.Enabled = False
    MDIForm1.mnuentry = False
    MDIForm1.mnusearch = False
    MDIForm1.mnuexit = False
End Sub

Private Sub mnuexit_Click()
a = MsgBox("Do You Want To Exit ?", vbQuestion + vbYesNo, "Diet Hub")
    If a = vbYes Then
        End
    End If
End Sub
