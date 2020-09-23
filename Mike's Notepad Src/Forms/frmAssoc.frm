VERSION 5.00
Begin VB.Form frmAssoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Default..."
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmAssoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicFocus 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   2760
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox ChSC 
         Caption         =   "Internet Explorer Source Code Viewer"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   3135
      End
      Begin VB.CheckBox ChSendto 
         Caption         =   "Place Professional NotePad On The Send To Menu..."
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Associate Windows Notepad with Plain Text Files..."
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Associate Professional NotePad  with Plain Text Files, and It's Own extension (*.pnt)..."
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ASloading As Boolean 'Are we loading the form ?
Private Sub ChSC_Click()
    CheckEnabled
End Sub
Private Sub ChSendto_Click()
    CheckEnabled
End Sub

Private Sub cmdApply_Click()
    'Do the job
    If Option1.Value Then AssociateText
    If Option2.Value Then AssociateNotepad
    If ChSC.Value = 1 Then
        AddSCviewer
        Else
        RemoveSCviewer
    End If
    If ChSendto.Value = 1 Then
        AddShortCutSendTo 'create shortcut
    Else
        If FileExists(SpecialFolder(9) + "\Michael J. Hardy's Professional NotePad.lnk") Then Kill SpecialFolder(9) + "\Michael J. Hardy's Professional NotePad.lnk"
    End If
  Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me 'bail
End Sub
Private Sub Form_Load()
    ASloading = True 'we're loading
    Me.Icon = frmMain.Icon 'same icon so why have two - use the same one
    'set control values appropriately
    Option1.Value = IsAssociatedText
    Option2.Value = IsNotePadAssociatedText
    ChSC.Value = IIf(IsSCviewer, 1, 0)
    ChSendto.Value = IIf(FileExists(SpecialFolder(9) + "\Michael J. Hardy's Professional NotePad.lnk"), 1, 0)
End Sub
Public Sub CheckEnabled()
    'compare with original states - if different enable 'Apply'
    cmdApply.Enabled = False
    If Option1.Value <> IsAssociatedText Then cmdApply.Enabled = True
    If Option2.Value <> IsNotePadAssociatedText Then cmdApply.Enabled = True
      If ChSC.Value <> IIf(IsSCviewer, 1, 0) Then cmdApply.Enabled = True
     If ChSendto.Value <> IIf(FileExists(SpecialFolder(9) + "\Michael J. Hardy's Professional NotePad.lnk"), 1, 0) Then cmdApply.Enabled = True
 'If PNTAssoc.Value <> IsAssociatedPNTText Then cmdApply.Enabled = True

End Sub
Private Sub Form_Paint()
    If ASloading Then 'OK we're loaded - do stuff we can only do once loaded
        PicFocus.SetFocus
        CheckEnabled
        ASloading = False 'Done it once - dont do it again
    End If
End Sub
Private Sub Option1_Click()
    CheckEnabled
End Sub
Private Sub Option2_Click()
    CheckEnabled
End Sub
