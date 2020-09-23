VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About This Awesome NotePad Software?"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   HelpContextID   =   1340
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Your System Information:"
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Alright Enough Already, Close This Screen..."
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   6480
      Width           =   3615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10610
      _Version        =   393216
      TabOrientation  =   1
      MousePointer    =   99
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      MouseIcon       =   "frmAbout.frx":058A
      TabCaption(0)   =   "About This?"
      TabPicture(0)   =   "frmAbout.frx":08A4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Awards Given!"
      TabPicture(1)   =   "frmAbout.frx":08C0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "More Credit's..."
      TabPicture(2)   =   "frmAbout.frx":08DC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label22"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "About The Author?"
      TabPicture(3)   =   "frmAbout.frx":08F8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Software License!"
      TabPicture(4)   =   "frmAbout.frx":0914
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   4215
         Left            =   -74040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "frmAbout.frx":0930
         Top             =   480
         Width           =   6495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   -74040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmAbout.frx":5103
         Top             =   1680
         Width           =   6615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Developer and Programmer:"
         Height          =   1815
         Left            =   3240
         TabIndex        =   9
         Top             =   1440
         Width           =   3375
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "   Postal Mail: 225 North Park Street,                Apartment 104, Sullivan Missouri,                            63080 USA..."
            Height          =   615
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - Michael J. Hardy -"
            Height          =   195
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email: notepad@fidmail.com"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   720
            MouseIcon       =   "frmAbout.frx":5723
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   720
            Width           =   1995
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Development Status:"
         Height          =   1095
         Left            =   3240
         TabIndex        =   8
         Top             =   3360
         Width           =   3375
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":5A2D
            Height          =   855
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5415
         Left            =   -73920
         TabIndex        =   3
         Top             =   60
         Width           =   6615
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   240
            Picture         =   "frmAbout.frx":5ABD
            ScaleHeight     =   645
            ScaleWidth      =   1875
            TabIndex        =   4
            Top             =   -960
            Width           =   1875
         End
         Begin VB.Image Image4 
            Height          =   1560
            Left            =   240
            Picture         =   "frmAbout.frx":747F
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label Label25 
            Caption         =   "Updated When New Version(s) of The Source Code are Available..."
            Height          =   495
            Left            =   1920
            TabIndex        =   19
            Top             =   4320
            Width           =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Updated When New Versions Are Released..."
            Height          =   375
            Left            =   1560
            TabIndex        =   18
            Top             =   2160
            Width           =   3255
         End
         Begin VB.Image Image3 
            Height          =   1290
            Left            =   2280
            Picture         =   "frmAbout.frx":8349
            Top             =   2640
            Width           =   1620
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "http://www.freewarefiles.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2040
            MouseIcon       =   "frmAbout.frx":9DB8
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1800
            Width           =   2100
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "http://www.planetsourcecode.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1920
            MouseIcon       =   "frmAbout.frx":A0C2
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3960
            Width           =   2490
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Thanks To All Who Have Voted For This Great NotePad Clone..."
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   840
            TabIndex        =   5
            Top             =   5040
            Width           =   4590
         End
         Begin VB.Image Image2 
            Height          =   1365
            Left            =   2400
            Picture         =   "frmAbout.frx":A3CC
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Special Thanks To The Following People and Companies:"
         Height          =   3855
         Left            =   -73920
         TabIndex        =   2
         Top             =   660
         Width           =   6375
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   2415
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   "frmAbout.frx":BB62
            Top             =   840
            Width           =   6135
         End
      End
      Begin VB.Image Image7 
         Height          =   4200
         Left            =   120
         Picture         =   "frmAbout.frx":BF35
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2520
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Professional NotePad!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   2550
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":DB9B
         Height          =   615
         Left            =   1440
         TabIndex        =   16
         Top             =   4800
         Width           =   6375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thank You For Using This SoftWare..."
         Height          =   195
         Left            =   -72360
         TabIndex        =   15
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Thank You For Using Professional NotePad... I hope that you will find this Software as Useful as I do..."
         Height          =   495
         Left            =   -73320
         TabIndex        =   14
         Top             =   4560
         Width           =   5295
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Written and Created By Michael J. Hardy Using Visual Basic 6.0 (Professional Edition!)"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   -73800
         TabIndex        =   13
         Top             =   5280
         Width           =   6375
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   2760
         Picture         =   "frmAbout.frx":DC9C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version x.x.xx"
         Height          =   555
         Left            =   3840
         TabIndex        =   12
         Top             =   840
         Width           =   2880
      End
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   120
      Picture         =   "frmAbout.frx":13B26
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1635
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   7560
      Picture         =   "frmAbout.frx":17391
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1635
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const SW_SHOW = 1
Const SW_SHOWMAXIMIZED = 3
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nshowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call StartSysInfo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Label9.Caption = "Released As Open Source Software..." & Chr(13) & "Under The General Public License!"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    Dim rc As Long
    Dim SysInfoPath As String
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  
   ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
    
        Else
            GoTo SysInfoErr
        End If
        Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento.", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Label2_Click()
'Unload Me
 
Dim lRet As Long
    Dim sText As String

    sText = "mailto:notepad@fidmail.com"
    lRet = ShellExecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
Unload Me
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Unable to start email client!"
    End If

End Sub

Private Sub Label3_Click()
Dim lRet As Long

    Dim sText As String

    sText = "http://www.freewarefiles.com"
    lRet = ShellExecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
    Unload Me
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Unable to your web browser!"
    End If

End Sub

Private Sub Label4_Click()

Dim lRet As Long

    Dim sText As String

    sText = "http://www.planetsourcecode.com"
    lRet = ShellExecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
Unload Me
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Unable to start your web browser!"
    End If
End Sub

