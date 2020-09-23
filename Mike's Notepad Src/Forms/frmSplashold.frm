VERSION 5.00
Begin VB.Form frmSplashOld 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   379
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   2640
      Top             =   2880
   End
End
Attribute VB_Name = "frmSplashOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim AboutLoc As String

    On Error Resume Next
    AboutLoc = App.Path & "\Skins\mjh.pnc"
    If Dir$(AboutLoc) <> "" Then
        Set Me.Picture = LoadPicture(AboutLoc)
    End If
End Sub

Private Sub Timer1_Timer()
Unload Me
frmMain.Show
End Sub
