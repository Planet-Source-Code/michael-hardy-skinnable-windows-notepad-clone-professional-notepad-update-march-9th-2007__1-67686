VERSION 5.00
Begin VB.Form Credits 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About This Awesome Text Editor ?"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6615
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close This Screen!"
      Height          =   495
      Left            =   1680
      Picture         =   "About.frx":058A
      TabIndex        =   0
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":2B97
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   2280
      Picture         =   "About.frx":2C42
      Top             =   2880
      Width           =   1920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Written and Created By Michael J. Hardy..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "™ Professional NotePad!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   120
      Picture         =   "About.frx":524F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1200
   End
   Begin VB.Menu cool 
      Caption         =   "cool"
      Visible         =   0   'False
      Begin VB.Menu coolmike 
         Caption         =   "mike is kool"
      End
   End
End
Attribute VB_Name = "Credits"
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
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Const COLOR_BACKGROUND = 2 'Colour of the background with no wallpaper
Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Const COLOR_MENU = 4 'Menu
Const COLOR_WINDOW = 5 'Windows background
Const COLOR_WINDOWFRAME = 6 'Window frame
Const COLOR_MENUTEXT = 7 'Window Text
Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Const COLOR_ACTIVEBORDER = 10 'Border of active window
Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Const COLOR_HIGHLIGHT = 13 'Selected item background
Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Const COLOR_BTNFACE = 15 'Button
Const COLOR_BTNSHADOW = 16 '3D shading of button
Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Const COLOR_BTNTEXT = 18 'Button text
Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color

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


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'    Dim i As Integer 'counter
 '    Dim AboutLoc As String
'
    On Error Resume Next
 '   AboutLoc = App.Path & "\Skins\mjh.pnc"
  '  If Dir$(AboutLoc) <> "" Then
   '     Set Me.Picture = LoadPicture(AboutLoc)
    'End If
GetSysColor WS_EX_TRANSPARENT
Me.DrawStyle = WS_EX_TRANSPARENT
End Sub

