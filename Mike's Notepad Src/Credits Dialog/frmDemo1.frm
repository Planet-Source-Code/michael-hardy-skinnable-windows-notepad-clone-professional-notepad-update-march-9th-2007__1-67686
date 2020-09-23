VERSION 5.00
Begin VB.Form Credits 
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemo1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmDemo1.frx":1CCA
   ScaleHeight     =   8100
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   2400
   End
   Begin VB.PictureBox picCredits 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5865
      Left            =   480
      ScaleHeight     =   5805
      ScaleWidth      =   9660
      TabIndex        =   0
      Top             =   1680
      Width           =   9720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Professional NotePad! (v2.0) - Beta # 13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   7560
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "frmDemo1.frx":1B91C
      Top             =   0
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here To Close!"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
'
'   SCROLLING AND FADING CREDITS
'
'   Written by Dirk Rijmenants (c) 2005
'------------------------------------------------------------

Option Explicit
 Private Declare Function WinExec Lib "Kernel32" _
        (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

  

Dim CreditLine()    As String
Dim CreditLeft()    As Long
Dim ColorFades(100) As Long
Dim ScrollSpeed     As Integer
Dim ColText         As Long
Dim FadeIn          As Long
Dim FadeOut         As Long

Dim cDiff1          As Long
Dim cDiff2          As Double
Dim cDiff3          As Double

Dim TotalLines      As Integer
Dim LinesOffset     As Integer
Dim Yscroll         As Long
Dim CharHeight      As Integer
Dim LinesVisible    As Integer

Private Sub Form_Load()
Dim FileO       As Integer
Dim fileName    As String
Dim tmp         As String
Dim i           As Integer

Dim Rcol1       As Long
Dim Gcol1       As Long
Dim Bcol1       As Long

Dim Rcol2       As Long
Dim Gcol2       As Long
Dim Bcol2       As Long

Dim Rfade       As Long
Dim Gfade       As Long
Dim Bfade       As Long

Dim PercentFade As Integer
Dim TimeInterval As Integer
Dim AlignText  As Integer

'################################################################
'
'   Preset the text and fade properties
'
'   To change text and background color or font you have to
'   change these properties on the picCredits picturebox
'   You can also use any image as background
'
    PercentFade = 30
    'The PercentFade sets the percentage of the box is used
    ' to fade in and out (set to zero when image used as background)
    TimeInterval = 45
    ScrollSpeed = 15
'   you might need to experiment with the ScrollSpeed and TimeInterval
    AlignText = 2 '( 1=left 2=center 3=right )
    
'################################################################

'set the number of line to be printed in the box
LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

'add empty lines at beginning to start off
For i = 1 To LinesVisible
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
Next

'Load the credits file
'  You could write the text also in the CreditLine(n) array,
'  but you than have to add the value TotalLines yourself!
FileO = FreeFile
fileName = App.Path & "\Skins\MJH.txt"
If Dir(fileName) = "" Then
        GoTo errHandler
        End If
On Error GoTo errHandler
Open fileName For Input As FileO
While Not EOF(FileO)
    Line Input #FileO, tmp
    ReDim Preserve CreditLine(TotalLines) As String
    CreditLine(TotalLines) = tmp
    TotalLines = TotalLines + 1
    Wend
Close #FileO

'set timer interval
Me.Timer1.Interval = TimeInterval

'set the number of line to be printed in the box
LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

'Next, we calculate a lot of time-eating stuff in advance.
'This is done before, to speedup timer sub ;-)

'set the fade-in and fade-out regions
CharHeight = picCredits.TextHeight("A")
If PercentFade <> 0 Then
    FadeOut = ((picCredits.Height / 100) * PercentFade) - CharHeight
    FadeIn = (picCredits.Height - FadeOut) - CharHeight - CharHeight
    Else
    FadeIn = picCredits.Height
    FadeOut = 0 - CharHeight
    End If
    
'set the percent values, ready for instant use later
ColText = picCredits.ForeColor
cDiff1 = (picCredits.Height - (CharHeight - 10)) - FadeIn
cDiff2 = 100 / cDiff1
cDiff3 = 100 / FadeOut

'calculate the left-position of each line, to center it
ReDim CreditLeft(TotalLines - 1)
For i = 0 To TotalLines - 1
    Select Case AlignText
    Case 1
        CreditLeft(i) = 100
    Case 2
        CreditLeft(i) = (picCredits.Width - picCredits.TextWidth(CreditLine(i))) / 2
    Case 3
        CreditLeft(i) = picCredits.Width - picCredits.TextWidth(CreditLine(i)) - 100
    End Select
Next i

'calculate 100 fade values from backcolor to forecolor
'(another time-eating thing done in advance)
Rcol1 = picCredits.ForeColor Mod 256
Gcol1 = (picCredits.ForeColor And vbGreen) / 256
Bcol1 = (picCredits.ForeColor And vbBlue) / 65536
Rcol2 = picCredits.BackColor Mod 256
Gcol2 = (picCredits.BackColor And vbGreen) / 256
Bcol2 = (picCredits.BackColor And vbBlue) / 65536
For i = 0 To 100
    Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
    Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
    Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
    ColorFades(i) = RGB(Rfade, Gfade, Bfade)
Next

'hit the throttle
Me.Timer1.Enabled = True
Exit Sub

errHandler:
Close FileO
MsgBox "Could not load Credits", vbCritical, " Credits Demo"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label1_Click(Index As Integer)
'WinExec("explorer /e,C:\\temp.txt,/select", SW_SHOW);
Unload Me
'Dim errReturn
 '     errReturn = WinExec("Data\mjh.com", 10)

End Sub

Private Sub picCredits_Click()
'Call ShellExecute(hwnd, "Open", "Data\mjh.com", App.Path, 1)
Unload Me
'Dim errReturn
 '     errReturn = WinExec("Data\mjh.com", 10)
End Sub

Private Sub Timer1_Timer()
Dim Ycurr       As Long
Dim TextLine    As Integer
Dim ColPrct     As Long
Dim i           As Integer
'clear pic for next draw
picCredits.Cls
Yscroll = Yscroll - ScrollSpeed
'calculate beginscroll
If Yscroll < (0 - CharHeight) Then
    Yscroll = 0
    LinesOffset = LinesOffset + 1
    If LinesOffset > TotalLines - 1 Then LinesOffset = 0
    'the offset sets the first line of the serie to be printed
    'this offset goes to the next line after each completely
    'scrolled line
    End If
'set Y for first  line
picCredits.CurrentY = Yscroll
Ycurr = Yscroll
'print only the visible lines
For i = 1 To LinesVisible
    If Ycurr > FadeIn And Ycurr < picCredits.Height Then
        'calculate fade-in forecolor
        ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    ElseIf Ycurr < FadeOut Then
        'calculate fade-out forecolor
        ColPrct = cDiff3 * Ycurr
        If ColPrct < 0 Then ColPrct = 0
        If ColPrct > 100 Then ColPrct = 100
        picCredits.ForeColor = ColorFades(ColPrct)
    Else
        'normal forecolor
        picCredits.ForeColor = vbRed 'ColText
    End If
    'get next line with offset
    TextLine = (i + LinesOffset) Mod TotalLines
    'set the X aligne value
    picCredits.CurrentX = CreditLeft(TextLine)
    'print that line
    picCredits.Print CreditLine(TextLine)
    'set Y to print next line
    Ycurr = Ycurr + CharHeight
Next i
End Sub


