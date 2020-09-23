VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmLoadWords 
   Caption         =   "Load Words into Database"
   ClientHeight    =   1485
   ClientLeft      =   9825
   ClientTop       =   4575
   ClientWidth     =   6585
   Icon            =   "frmLoadWords.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Words"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5490
      TabIndex        =   3
      Top             =   675
      Width           =   1050
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6015
      TabIndex        =   2
      Top             =   240
      Width           =   390
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   5850
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   1185
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11086
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLoadWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
On Error Resume Next
    LoadWords txtFileName.Text
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    With cmndlg
        .filefilter = "Plain text (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFile
        If Len(.filename) = 0 Then Exit Sub
        
        txtFileName.Text = .filename
    End With
End Sub

Private Sub Form_Load()
On Error Resume Next
    'LoadWords ""
End Sub

