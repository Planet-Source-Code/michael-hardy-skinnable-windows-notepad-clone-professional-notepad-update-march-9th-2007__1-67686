VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Professional NotePad By Michael J. Hardy"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   1560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This should provide some decent reference for
'creating a splash screen that fades in slowly
'waits a few seconds, and then fades out slowly.
'I give credit to some VB students that helped me out with
'the loops.



'API Declarations
'*****************************************

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long


Private Declare Function SetLayeredWindowAttributes Lib "user32" _
(ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As Long

'Variables
'******************************************
'constants that work with API from above
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
'basic variables
Dim i As Integer
Dim sdelay As String
Dim tdelay As Single


Private Sub Form_Load()
    Dim i As Integer 'counter
     Dim AboutLoc As String

    On Error Resume Next
    AboutLoc = App.Path & "\Skins\mjh.pnc"
    If Dir$(AboutLoc) <> "" Then
        Set Me.Picture = LoadPicture(AboutLoc)
    End If
'Get attributes of Splash Screen
 '  SetWindowLong frmSplash.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
  '  SetLayeredWindowAttributes frmSplash.hwnd, 0, 1, LWA_ALPHA
'set FadeIn timer interval
  '  tmrFISplash.Interval = 200
'Set heading label caption
  ' lblHeading.Caption = "Your Program Name Here"
'set version caption
    'l'blVersion.Caption = "Version: " & " " & App.Major & "." & App.Minor

    
End Sub

Public Sub FadeIn()

'Set delay for timer to a very low number - determines load speed
'make smaller number to make form load faster
   ' sdelay = 0.00000013
'loop for fading in window
   ' For i = 0 To 255
'sets windows visibility attributes
    '    SetLayeredWindowAttributes frmSplash.hwnd, 0, i, LWA_ALPHA
'increase timer interval
     '   tdelay = Timer + Val(sdelay)
'let Windows do it's thing
      '  While tdelay > Timer: DoEvents: Wend
'increase i by 1
    'Next i
'set fade out timer interval
   ' tmrFOSplash.Interval = 600
    
End Sub

Public Sub FadeOut()

'Set delay for timer to a very low number - determines load speed
'make smaller number to make form load faster
    'sdelay = 0.00001
'loop for fading out window
    'For i = 255 To 0 Step -1
'sets window visibility attributes
     '   SetLayeredWindowAttributes frmSplash.hwnd, 0, i, LWA_ALPHA
'increase timer interval
      ' tdelay = Timer + Val(sdelay)
'let Windows do its thing
       'While tdelay > Timer: DoEvents: Wend
'decrease i by 1
    'Next i
'unload splash form from memory
    'Unload Me
'frmlogin.show
'enter code to do events you want
'frmMain.Show

    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call FadeOut(Me, True)
    Unload Me   'unload from memory
    'you would do this so users can skip the splash screen
    'by clicking it
    frmMain.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Call FadeOut
frmMain.Show
End Sub

Private Sub tmrFISplash_Timer()
 '   Call FadeIn                     'calls Sub
  '  tmrFISplash.Enabled = False     'disables FadeIn timer
End Sub

Private Sub tmrFOSplash_Timer()
   ' Call FadeOut                    'calls sub
    'tmrFOSplash.Enabled = False     'disables FadeOut timer
    
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
