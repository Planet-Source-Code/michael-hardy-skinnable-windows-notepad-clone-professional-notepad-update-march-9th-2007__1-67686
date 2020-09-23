VERSION 5.00
Begin VB.Form frmSpelling 
   Caption         =   "Spell Checker... - Check Your Spelling!"
   ClientHeight    =   3750
   ClientLeft      =   11475
   ClientTop       =   4470
   ClientWidth     =   5715
   Icon            =   "frmSpelling.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReplaceWith 
      Height          =   285
      Left            =   75
      TabIndex        =   9
      Top             =   765
      Width           =   4185
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   315
      Left            =   4290
      TabIndex        =   8
      Top             =   1665
      Width           =   1410
   End
   Begin VB.CommandButton cmdIgnoreAll 
      Caption         =   "Ignore All"
      Height          =   285
      Left            =   4290
      TabIndex        =   7
      Top             =   2745
      Width           =   1410
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore"
      Height          =   300
      Left            =   4290
      TabIndex        =   6
      Top             =   2385
      Width           =   1410
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   315
      Left            =   4290
      TabIndex        =   4
      Top             =   1305
      Width           =   1410
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   300
      Left            =   4290
      TabIndex        =   3
      Top             =   2040
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4290
      TabIndex        =   2
      Top             =   3360
      Width           =   1410
   End
   Begin VB.ListBox lstMatches 
      Height          =   2385
      IntegralHeight  =   0   'False
      Left            =   75
      TabIndex        =   1
      Top             =   1305
      Width           =   4170
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   210
      Width           =   4185
   End
   Begin VB.Label Label3 
      Caption         =   "Replace"
      Height          =   240
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "With"
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Suggestions"
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   1035
      Width           =   2055
   End
End
Attribute VB_Name = "frmSpelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    AddWord txtWord.Text, False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    bReplaceText = False
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    Unload Me
End Sub

Private Sub cmdIgnoreAll_Click()
    ReDim Preserve sTempWords(UBound(sTempWords) + 1)
    sTempWords(UBound(sTempWords)) = txtWord.Text
    Unload Me
End Sub

Private Sub cmdReplace_Click()
    If Len(txtReplaceWith.Text) > 0 Then
        sTextBeingChecked = Replace(sTextBeingChecked, txtWord.Text, txtReplaceWith.Text, 1, 1)
        bReplaceText = True
    End If
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    If Len(txtReplaceWith.Text) > 0 Then
        sTextBeingChecked = Replace(sTextBeingChecked, txtWord.Text, txtReplaceWith.Text, 1)
        ChangeReplaced txtWord.Text, txtReplaceWith.Text
        bReplaceText = True
    End If
    Unload Me
End Sub

Private Sub lstMatches_Click()
    txtReplaceWith.Text = lstMatches.List(lstMatches.ListIndex)
End Sub
