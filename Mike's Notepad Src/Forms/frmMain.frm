VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Object = "{52BD8A52-B792-4C45-A4D9-245CC945AC34}#1.0#0"; "wbocx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Michael J. Hardy's Professional NotePad!"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   22
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "New Document"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Open Document..."
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Save Document..."
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Save The Document As..."
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Copy..."
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Paste Into Document..."
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Delete Text"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Select All"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Undo..."
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Redo..."
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Find within the Document..."
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Check Your Spelling..."
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Print Document..."
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Choose A Different Theme...."
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "About This Cool Text Editor?"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1800
      Top             =   960
   End
   Begin MSComDlg.CommonDialog CDl 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer PasteTimer 
      Interval        =   100
      Left            =   6600
      Top             =   1440
   End
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   0
      ScaleHeight     =   4890
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   510
      Width           =   1215
      Begin RichTextLib.RichTextBox RTF 
         Height          =   735
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1296
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"frmMain.frx":6764
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5847
            Picture         =   "frmMain.frx":67E7
            Text            =   "Mike's Professional NotePad Is Ready!"
            TextSave        =   "Mike's Professional NotePad Is Ready!"
            Object.ToolTipText     =   "File path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Cursor position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Selection length"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "2 bytes"
            TextSave        =   "2 bytes"
            Object.ToolTipText     =   "File size"
         EndProperty
      EndProperty
      MouseIcon       =   "frmMain.frx":6E10
   End
   Begin RichTextLib.RichTextBox RTFtemp 
      DragIcon        =   "frmMain.frx":7CEA
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":7E3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList4 
      Left            =   6600
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8F51
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9FE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":B075
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":C107
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D199
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E22B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":F2BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1034F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":113E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":12473
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13505
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14597
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":15629
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":166BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16C55
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   5760
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":17AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":18B39
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19BCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1AC5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1BCEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1CD81
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DE13
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1EEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1FF37
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20FC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2205B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":230ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2417F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":25211
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":262A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2683D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   5040
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2768F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":28721
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":297B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A845
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2B8D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2C969
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2D9FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2EA8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2FB1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":30BB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":31C43
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":32CD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":33D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":34DF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":35E8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":36425
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin WBOCXLib.Wbocx Wb 
      Left            =   1320
      Top             =   360
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Professional NotePad!"
      End
      Begin VB.Menu mnusep13131313 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutPopup 
         Caption         =   "About This Awesome NotePad Software?"
      End
      Begin VB.Menu mnuExit13 
         Caption         =   "Exit This Great Application..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAssociations 
         Caption         =   "File Associations (Make Default Text Editor!)"
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Above"
         Index           =   9
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select Below"
         Index           =   10
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Date Stamp"
         Index           =   12
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find and Replace"
         Index           =   13
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Spell Check (Check Your Spelling)..."
         Index           =   15
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Load Spell Check Word List..."
         Index           =   16
      End
      Begin VB.Menu mnusep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupAbout 
         Caption         =   "&About Professional NotePad?"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatWordwrap 
         Caption         =   "Use Word Wrap (Recommended)..."
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFormatSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatFont 
         Caption         =   "Change The Font and Font Color!"
      End
      Begin VB.Menu mnuFormatBackcolor 
         Caption         =   "Change The Background Color..."
      End
      Begin VB.Menu mnuFormatSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatStats 
         Caption         =   "Document Statistics"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Click Here To Toggle The Toolbar!"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "View Statusbar... - (Click Here To Toggle!)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeperato 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Minimize To System Tray..."
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSepFridayThe13th 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplashScreen 
         Caption         =   "Show The Splash Screen At Startup..."
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mSkin 
      Caption         =   "&Theme..."
      Begin VB.Menu mnuSkin 
         Caption         =   "Choose A Different Theme..."
      End
   End
   Begin VB.Menu mnuSound 
      Caption         =   "&Sound"
      Begin VB.Menu mnuPlay 
         Caption         =   "Use Sound Effects... - (Click Here To Toggle!)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPlay2 
         Caption         =   "Listen to Soothing Music While You Type... (Click To Toggle!)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Infor&mation..."
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About This Awesome Text Editor?"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Professional NotePad!
'
' Copyright © January of 2007-2013...
'
' ® All Rights Are Reserved...
'
' Written, Created, and Designed By Michael J. Hardy
'
' Released Under The GNU (General Public License) (GPL)...
'
' YOU CAN NOT DISTRIBUTE OR PUBLISH THIS SOURCE CODE...
' YOU ALSO CAN NOT SELL ANY PORTION OF THIS SOFTWARE
' OR IT'S SOURCE CODE... ANY VIOLATION WILL RESULT IN
' IMMEDIATE TERMINATION OF THE SOFTWARE LICENSE AND
' WILL RESULT IN SEVERE CRIMINAL OR CIVIL ACTION...
' USE THIS SOURCE CODE FOR EDUCATIONAL PURPOSES ONLY!
' PLEASE CONSIDER HELPING WITH THIS DEVELOPMENT AND
' HELP ME TURN THIS SIMPLE NOTEPAD INTO A REAL ADVANCED
' WINDOWS APPLICATION... - THANK YOU - SINCERELY, -
' - MICHAEL J. HARDY - CREATIVE-CODING ® PRESIDENT! -
'
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
            (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
            ByVal lpsz2 As String) As Long

' Toolbar constants
Private Const WM_USER = &H400
Private Const TBSTYLE_FLAT As Long = &H800
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TB_SETIMAGELIST = WM_USER + 48
Private Const TB_SETHOTIMAGELIST = WM_USER + 52
Private Const TB_SETDISABLEDIMAGELIST = WM_USER + 54
Public WithEvents Undo As clsUndo 'heavily modified version of a class by Sebastian Thomschke
Attribute Undo.VB_VarHelpID = -1
Dim onlyLoading As Boolean 'indicates form load complete
Dim mTStop() As Boolean 'allows the use of 'tab' within the richtextbox
'rather than moving focus to the next control
Dim myCommand As String
 Const SKIN_DEFAULT As String = "FauxS-TOON\FauxS-TOON.uis" 'or whatever...
  Private Declare Function ShellAbout Lib "shell32.dll" Alias _
    "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As _
    String, ByVal szOtherStuff As String, ByVal hIcon As _
    Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000

    Private m_SkinName As String
    Private m_SkinPath As String

Private Sub Form_Activate()
 InitTB
End Sub

Private Sub Form_Initialize()
 InitCommonControls
' If mnuSplashScreen.Checked = True Then frmSplash.Show
   If mnuSplashScreen.Checked = True Then
If mnuSplashScreen.Checked = False Then
mnuSplashScreen.Caption = "Splash Screen Is Disabled... - (Click Here to Enable)..."
Else
End If
If mnuSplashScreen.Checked = True Then
mnuSplashScreen.Caption = "Splash Screen Is Enabled... - (Click Here To Disable The Splash Screen)"
frmMain.Hide

frmSplash.Show
If mnuPlay.Checked = True Then
PlaySound 104
End If
End If
End If
  If mnuSystray.Checked = True Then
'SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, _
 '       GWL_EXSTYLE) And Not WS_EX_APPWINDOW)

  With nid
       .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Michael J. Hardy's Skinnable Windows NotePad!" & vbNullChar
       
  Shell_NotifyIcon NIM_ADD, nid
  End With
  End If
 
'Else
'mnuSystray.Checked = True
' With nid
 '       .cbSize = Len(nid)
  '      .hWnd = Me.hWnd
   '     .uId = vbNull
    '    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     '   .uCallBackMessage = WM_MOUSEMOVE
      '  .hIcon = Me.Icon
       ' .szTip = "Your ToolTip" & vbNullChar
       '
       'Shell_NotifyIcon NIM_ADD, nid
       'End With
    'End If
'If mnuSystray.Checked = False Then
 'Shell_NotifyIcon NIM_DELETE, nid
'mnuSystray.Caption = "Minimize To System Tray..."
'Else
'mnuSystray.Caption = "Minimize To System Tray..."
'End If
If mnuPlay2.Checked = True Then
 gHW = Me.hwnd
    gsMidiFile = App.Path & "\Music\Music.mid"
    gbRepeat = True  'Set to true if you want it to repeat
    
    'You can specify the start and stop times with
    'these variables.  To Play the whole song, just
    'set both glFrom and glTo to 0.
    glFrom = 0
    glTo = 0 'Will play for a couple of seconds
    
    'Go ahead and make sure no other midi file is open
    CloseMidi
    UnHook
    DoEvents
    
    'It is important to hook the form after
    'you begin playing the Midi.  Not sure why.
    PlayMidi
    Hook
    End If
End Sub
Private Function InitTB(Optional intSize As Integer = 24) As Long

    ' Set up the toolbar
    Dim lngStyle As Long
    Dim lRes As Long
    
    ' Get the toolbar handle (we cannot just use tbrMain.hwnd as this is a container
    ' window for the actual toolbar control)
    Dim hTBar As Long
    hTBar = FindWindowEx(tbrMain.hwnd, 0&, "ToolbarWindow32", vbNullString)
        
    ' The style "TBSTYLE_FLAT" needs to be added.  Although this option is available
    ' in the property pages for the toolbar, it needs to be set here.
    
    ' Get the current style
    lngStyle = SendMessage(hTBar, TB_GETSTYLE, 0&, ByVal 0&)
    
    ' Add the TBSTYLE_FLAT style (could also apply other styles here)
    lngStyle = lngStyle Or TBSTYLE_FLAT
        
    ' Set the new style
    Call SendMessage(hTBar, TB_SETSTYLE, 0&, ByVal lngStyle)
    tbrMain.Refresh
    
    ' Now add the ImageList's for the normal, hot, and disabled states
    lRes = SendMessage(hTBar, TB_SETIMAGELIST, 0, ByVal ImageList2.hImageList)
    lRes = SendMessage(hTBar, TB_SETHOTIMAGELIST, 0, ByVal ImageList3.hImageList)
    lRes = SendMessage(hTBar, TB_SETDISABLEDIMAGELIST, 0, ByVal ImageList4.hImageList)
    
End Function

Private Sub Exit_Click()
    mnuFileExit_Click
End Sub
Private Sub About_Click()
    mnuFileAbout_Click
End Sub
Private Sub Form_Load()
    'get settings from registry
    
    mnuViewToolbar.Checked = True
    mnuPlay2.Checked = False
    mnuSystray.Checked = True
   ' mnuPlay2.Enabled = False
   ' mnuPopupAbout.Visible = False
    onlyLoading = True
    Wb.AddAdditionalThread (hwnd)
    Wb.ForceBackgrounds (hwnd)
     SetWindowLong Me.hwnd, GWL_EXSTYLE, (GetWindowLong(hwnd, _
        GWL_EXSTYLE) Or WS_EX_APPWINDOW)

  'frmMain.Show
      
  Const SKIN_DEFAULT As String = "VistaXp\VistaXp.uis" 'or whatever...
  
'Current path and file name of skin.
'Start the form with the saved skin from the registry or a default skin.
  
    Dim sName As String
       Wb.InitWB
    'Get the default skin.
        m_SkinName = SKIN_DEFAULT
        m_SkinPath = App.Path & "\Skins\" 'or the path to the skins file. (Could also be in the registry).
      
    'Get the saved skin.
        'sName = GetSetting ("Michael-J-Hardy's-Professional-Software\ + App.EXEName, "Settings", "Choose A Different Theme...", sName)
    sName = GetSetting("MichaelsKoolNotePadXp\" + App.EXEName, "Settings", "Choose A Different Theme...", sName)
    'Check if a skin file was found.
        If LenB(sName) Then
              
            'Set the current skin name for the module.
                m_SkinName = sName
                  
        End If
          
    'Load the current skin file here.
        Wb.LoadUIS (m_SkinName)
' We want to use the samples path
Wb.SetRootPath 1
'Dim skinpath As String
'VB.App.Path & "\skins\"
'skinpath = VB.App.Path & "\skins\"
Wb.SetRootPathStr m_SkinPath
Wb.SkinAllThreads
Wb.EnableSpecialBackgroundSupport (Me.hwnd)

' Specify a skin to load - note that this is a relative path from the root path
'Wb.LoadUIS m_SkinName '"FauxS-TOON\FauxS-TOON.uis"
Wb.UseToolbarBackground (Me.hwnd)
Wb.EnableStatusSkin
' Tell DirectSkin to start skinning this thread
Wb.DoWindow Me.hwnd
 
'    m_SkinName = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Skin", m_SkinName)
    ' Dim currWindowState As Integer

    ' in case no value is in the registry
    On Error Resume Next
    ' If the form is currently maximized or minimized, temporarily
    ' revert to normal state, otherwise the Move command fails.
   ' currWindowState = Me.WindowState
  '  If currWindowState <> 0 Then Me.WindowState = 0
    'Me.Wb = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.ExeName, "Settings", "Skin", filename)
      ' Use a Move method to avoid multiple Resize and Paint events.
'   Me.Left = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainLeft", 1000)
 '  Me.Top = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainTop", 1000)
'Me.Width = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainWidth", 8500)
 '  Me.Height = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainHeight", 6500)
 'General Window Settings: (Take as sample)
LoadRegistryValues
    myCommand = Command()
    InitCmnDlg Me.hwnd
    cmndlg.Flags = 5
    RTF.BackColor = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Backcolor", Str(vbWhite)))
    mnuSplashScreen.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "SplashScreen", True)
   
    mnuViewStatusbar.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Statusbar", True)
    mnuViewToolbar.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Click Here To Toggle The Toolbar!", True)
    mnuFormatWordwrap.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Wordwrap", True)
    mnuPlay.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Use  Sound Effects... - (Click Here To Toggle!)", True)
     mnuPlay2.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Listen to Soothing Music While You Type... (Click To Toggle!)", False)
   mnuSystray.Checked = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Minimize To System Tray...", mnuSystray)
    SB.Visible = mnuViewStatusbar.Checked
    tbrMain.Visible = mnuViewToolbar.Checked
    'mnuPlay2EffectEffect.Visible = mnuPlay2EffectEffect.Checked

    'frmSplash.Visible = mnuSplashScreen.Checked
    RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
    RTF.Text = " "
    RTF.SelStart = 0
    RTF.SelLength = 1
    SelFont
    RTF.Text = ""
    Set Undo = New clsUndo
    Undo.RichBox = RTF
    Undo.Reset
    FileChanged = False
 '   Wb.InitWB

' We want to use the samples path
'Wb.SetRootPath 1
'Dim skinpath As String
'skinpath = VB.App.Path & "\skins\"
'Wb.SetRootPathStr skinpath

' Specify a skin to load - note that this is a relative path from the root path
'Wb.LoadUIS "FauxS-TOON\FauxS-TOON.uis"
'Wb.ReloadUIS
'Wb.UseToolbarBackground (Me.hwnd)
'Wb.SkinAllThreads
'Wb.EnableStatusSkin
' Tell DirectSkin to start skinning this thread
'Wb.DoWindow Me.hwnd
'mnuPlay2EffectEffect.Checked = Not mnuPlay2EffectEffect.Checked
  '  mnuPlay2Effect.Checked = Not mnuPlay2Effect.Checked
'Set IconObject = frmMain.Icon
 '   AddIcon frmMain, IconObject.Handle, IconObject, "Michael J. Hardy's Professional NotePad!"

End Sub
Private Sub LoadRegistryValues()
 'General Window Settings: (Take as sample)
 With Me
  .Width = GetSetting(App.EXEName, "WindowInfo\WindowSizes\Width", .Name, .Width)
  .Height = GetSetting(App.EXEName, "WindowInfo\WindowSizes\Height", .Name, .Height)
  
  .Left = GetSetting(App.EXEName, "WindowInfo\WindowPositions\Left", .Name, .Left)
  .Top = GetSetting(App.EXEName, "WindowInfo\WindowPositions\Top", .Name, .Top)
    
  .WindowState = GetSetting(App.EXEName, "WindowInfo\WindowStates", .Name, .WindowState)
 End With
 'End of General Window Settings
 
 'Misc Settings: (Enter code here)
 'End of Misc Settings
End Sub
Private Sub SaveRegistryValues()
 'General Window Settings: (Take as sample)
 With Me
  SaveSetting App.EXEName, "WindowInfo\WindowStates", .Name, .WindowState
    
  Dim colSize As Collection
  Set colSize = GetPosAndSize(Me)
    
  SaveSetting App.EXEName, "WindowInfo\WindowSizes\Width", .Name, colSize.Item("Width")
  SaveSetting App.EXEName, "WindowInfo\WindowSizes\Height", .Name, colSize.Item("Height")
  
  SaveSetting App.EXEName, "WindowInfo\WindowPositions\Left", .Name, colSize.Item("Left")
  SaveSetting App.EXEName, "WindowInfo\WindowPositions\Top", .Name, colSize.Item("Top")
  
  Set colSize = Nothing
 End With
 'End of General Window Settings
 
 'Misc Settings: (Enter code here)
 'End of Misc Settings
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
      Dim Msg As Long
       'the value of X will vary depending
       'upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        Msg = X
       Else
        Msg = X / Screen.TwipsPerPixelX
       End If
       Select Case Msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopupSys
       End Select
End Sub

Private Sub Form_Paint()
    If onlyLoading Then
        If myCommand <> "" Then
            'We've been shelled
            DoEvents
            NoStatusUpdate = True
            Screen.MousePointer = 11
            SB.Panels(1) = "Loading file...."
            LockWindowUpdate Me.hwnd
            myCommand = strUnQuoteString(myCommand) 'sometimes explorer uses quotes('send to' for example)
            myCommand = GetLongFilename(myCommand) 'looks better than a dos path
            Select Case LCase(ExtOnly(myCommand))
                Case "txt"
                    RTF.SelText = OneGulp(myCommand) 'binary read
                Case "rtf"
                    RTFtemp.LoadFile myCommand 'rtf load
                    RTF.SelText = RTFtemp.Text
                Case "doc"
                    OpenWordDoc myCommand 'see sub - returns plain text
                Case Else
                    RTF.SelText = OneGulp(myCommand) 'otherwise do binary read
            End Select
            Me.Caption = "Now Viewing - " + FileOnly(myCommand)
            SB.Panels(1) = myCommand
            SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
            RTF.Tag = myCommand
            If FileLen(myCommand) > 100000 Then
                'just using SelFont, RTF selection falls
                'over somewhere around 100k so do this
                'slightly less efficient but more reliable
                'method of font control
                RTF.SelStart = 0
                RTF.SelLength = Len(RTF.Text)
                SelFont
                RTF.SelLength = 0
            End If
            NoStatusUpdate = False
            EditEnable
            RTF.SelStart = 0
            Screen.MousePointer = 0
            LockWindowUpdate 0
        End If
        Undo.Reset
        FileChanged = False
        onlyLoading = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
     If mnuPlay.Checked = True Then
PlaySound 107
End If
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbQuestion & vbYesNoCancel, "Michael's Professional NotePad!")
       
        Select Case Response
            Case vbCancel
                Cancel = 1 'dont unload, user must want to do something else after all
          Case vbYes
          mnuFileSave_Click
                'SaveAFile returns false if user cancels during save process - dont unload
                'If Not SaveAFile Then Cancel = 1
        End Select
    End If
    SaveRegistryValues
    'save settings to registry
    'SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Skin", m_SkinName
  ' SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Choose A Different Theme...", m_SkinName
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Use  Sound Effects... - (Click Here To Toggle!)", mnuPlay.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Listen to Soothing Music While You Type... (Click To Toggle!)", mnuPlay2.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Minimize To System Tray...", mnuSystray.Checked
    
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Wordwrap", mnuFormatWordwrap.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Statusbar", mnuViewStatusbar.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "SplashScreen", mnuSplashScreen.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Click Here To Toggle The Toolbar!", mnuViewToolbar.Checked
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainLeft", Me.Left
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainTop", Me.Top
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainWidth", Me.Width
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "MainHeight", Me.Height
     'SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "skin", m_SkinName
       'delIcon IconObject.Handle
    'delIcon frmMain.Icon.Handle
    If mnuSystray.Checked = True Then
    Shell_NotifyIcon NIM_DELETE, nid
    End If
      End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'If WindowState.maximized Then WindowState
    'placing the RTF on a left aligned picturebox makes
    'resizing easier
    PicLeft.Width = Me.Width
    If mnuSystray.Checked = True Then
    If Me.WindowState = vbMinimized Then Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'unload correctly
'      SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Skin", m_SkinName

      Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
       '
 ' delIcon IconObject.Handle
  '  delIcon frmMain.Icon.Handle
    Next
     Shell_NotifyIcon NIM_DELETE, nid
'Shell_NotifyIcon NIM_DELETE, nid
      End Sub

Private Sub mnuAboutPopup_Click()
On Error Resume Next
  '  If mnuPlay.Checked = True Then
'PlaySound 103
'End If
     'ShellAbout Me.hWnd, "Professional NotePad! - (Windows Notepad!)", _
     '  "Developed For ® Microsoft By Michael J. Hardy..." & "       Special Thanks To My Daughter (Zoé Hardy) - (© 2007)", Me.Icon
   
    'MsgBox "                        Professional NotePad!" + vbCrLf + vbCrLf + vbCrLf + "- This Software Was Written and Created By Michael J. Hardy..." + vbCrLf + " - Copyright © January of 2007..." + vbCrLf + "- Special Thanks To: My Wife (Kara), My Daughter (Zoe)" + vbCrLf + "  and My Parents (Jim and Sher Hardy)..." + vbCrLf + vbCrLf + "This Version Of Professional NotePad Is Freeware!" + vbCrLf + vbCrLf + "Please visit http://www.mikes-games.com/notepad/ For Updates...", vbInformation, "About This Awesome SoftWare?"
'Credits.Show]
If mnuPlay.Checked = True Then
PlaySound 103
End If
frmAbout.Show
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Dim curvl As Long, st As Long
    
    LockWindowUpdate Me.hwnd
    Select Case Index
        Case 0 'Undo
            Undo.Undo
        Case 1 'Redo
            Undo.Redo
        Case 3 'cut
            Undo.Cut
        Case 4 'copy
            Undo.Copy
        Case 5 'paste
            Undo.Paste
        Case 6 'delete
            Undo.Delete
        Case 8 'select all
            RTF.SelStart = 0
            RTF.SelLength = Len(RTF.Text)
            RTF.SetFocus
        Case 9 'select above - maintaining scroll position
            Screen.MousePointer = 11
            'this is where the top line and therefore scroll position is
            curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            'do the selection
            st = RTF.SelStart
            RTF.SelStart = 0
            RTF.SelLength = st
            'return the scroll postion back to what it was - see SetScrollPos sub
            SetScrollPos curvl, RTF
            RTF.SetFocus
        Case 10 'select below
            Screen.MousePointer = 11
            curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            st = RTF.SelStart + RTF.SelLength
            RTF.SelStart = st
            RTF.SelLength = Len(RTF.Text) - st
            SetScrollPos curvl, RTF
            RTF.SetFocus
        Case 12 'date stamp
            frmDate.Show vbModal, Me
        Case 13 'find
            frmFind.Show , Me
        Case 15 'Spell check
        On Error Resume Next
            If RTF.SelLength <> 0 Then
                ' just spell check the highlighted text
                On Error Resume Next
                RTF.TextRTF = Mid$(RTF.Text, RTF.SelStart, RTF.SelLength)
            Else
                ' Spell check everything
                On Error Resume Next
                LoadWords App.Path & "\Spell Checker\Wl.txt"
                RTF.Text = SpellCheck(RTF.Text)
            End If
        Case 16 ' Load word list
            frmLoadWords.Show vbModal, Me
    End Select
    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Private Sub mnuExit13_Click()
mnuFileExit_Click
End Sub

Private Sub mnuFileAbout_Click()
    'blah blah
    On Error Resume Next
  '  If mnuPlay.Checked = True Then
'PlaySound 103
'End If
     ShellAbout Me.hwnd, "Professional NotePad! - (Windows Notepad!)", _
       "Developed For ® Microsoft By Michael J. Hardy..." & "       Special Thanks To My Daughter (Zoé Hardy) - (© 2007)", Me.Icon
   
    'MsgBox "                        Professional NotePad!" + vbCrLf + vbCrLf + vbCrLf + "- This Software Was Written and Created By Michael J. Hardy..." + vbCrLf + " - Copyright © January of 2007..." + vbCrLf + "- Special Thanks To: My Wife (Kara), My Daughter (Zoe)" + vbCrLf + "  and My Parents (Jim and Sher Hardy)..." + vbCrLf + vbCrLf + "This Version Of Professional NotePad Is Freeware!" + vbCrLf + vbCrLf + "Please visit http://www.mikes-games.com/notepad/ For Updates...", vbInformation, "About This Awesome SoftWare?"
'Credits.Show]
If mnuPlay.Checked = True Then
PlaySound 103
End If
frmAbout.Show
End Sub

Private Sub mnuFileAssociations_Click()
    frmAssoc.Show vbModal, Me
End Sub
Private Sub mnuFileExit_Click()
     Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
    End If
     If mnuPlay.Checked = True Then
PlaySound 107
End If
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbQuestion & vbYesNoCancel, "Michael's Professional NotePad!")
        
        Select Case Response
            Case vbCancel
                Exit Sub
            Case vbYes
          ' If Not SaveAFile Then
           '             Effect = vbDropEffectNone
            '            Exit Sub
             '       End If
    mnuFileSave_Click
            'End Select
            ' If Not SaveAFile Then Exit Sub
       End Select
   'End If
    'clear everything
    RTF.Text = ""
    Me.Caption = "Untitled" + " - Michael J. Hardy's Professional NotePad!"
    SB.Panels(1) = "This File Is Not Saved..."
    SB.Panels(4) = GetFileSize(2)
    RTF.Tag = ""
    SelFont 'maintain control of fonts
    Undo.Reset 'clear undo buffer
    FileChanged = False
    SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Skin", m_SkinName

    Unload Me
    End Sub
Private Sub mnuFileNew_Click()
    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
    'End If
     If mnuPlay.Checked = True Then
PlaySound 107
End If
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbQuestion & vbYesNoCancel, "Michael's Professional NotePad!")
        
        Select Case Response
            Case vbCancel
                Exit Sub
            Case vbYes
          ' If Not SaveAFile Then
           '             Effect = vbDropEffectNone
            '            Exit Sub
             '       End If
    mnuFileSave_Click
            'End Select
            ' If Not SaveAFile Then Exit Sub
       End Select
   End If
    'clear everything
    RTF.Text = ""
    Me.Caption = "Untitled " + " - Michael J. Hardy's Professional NotePad!"
    SB.Panels(1) = "This File Is Not Saved..."
    SB.Panels(4) = GetFileSize(2)
    RTF.Tag = ""
    SelFont 'maintain control of fonts
    Undo.Reset 'clear undo buffer
    FileChanged = False
    
End Sub
Private Sub mnuFileOpen_Click()

    Dim Response As VbMsgBoxResult
    If FileChanged Then 'do we save current doc ?
     If mnuPlay.Checked = True Then
PlaySound 107
End If
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbQuestion & vbYesNoCancel, "Michael's Professional NotePad!")
       
        Select Case Response
            Case vbCancel
                Exit Sub
            Case vbYes
          mnuFileSave_Click
 'If Not SaveAFile Then Exit Sub
        End Select
        End If
        
    With cmndlg
        .filefilter = "Plain Text (*.txt)|*.txt|Rich Text (*.rtf)|*.rtf|Professional Notepad Text (*.pnt)|*.pnt|All Files (*.*)|*.*"
        OpenFile
        If Len(.filename) = 0 Then Exit Sub
        SB.Panels(1) = "Loading file...."
        NoStatusUpdate = True
        Me.Refresh
        Screen.MousePointer = 11 'hourglass
        LockWindowUpdate Me.hwnd 'RTF.hWnd
        RTF.Text = ""
        SelFont
        Select Case .filefilterindex
            Case 1
                RTF.SelText = OneGulp(.filename) 'binary read
            Case 2
                RTFtemp.LoadFile .filename 'rtf load
                RTF.SelText = RTFtemp.Text
            Case 3
               RTFtemp.LoadFile .filename 'rtf load
                RTF.SelText = RTFtemp.Text
               ' OpenWordDoc .filename 'see sub - returns plain text
            Case 4
                RTF.SelText = OneGulp(.filename) 'otherwise do binary read
        End Select
        If FileLen(.filename) > 100000 Then
            'just using SelFont, RTF selection falls
            'over somewhere around 100k so do this
            'slightly less efficient but more reliable
            'method of font control
            RTF.SelStart = 0
            RTF.SelLength = Len(RTF.Text)
            SelFont
            RTF.SelLength = 0
        End If
         If mnuPlay.Checked = True Then
PlaySound 110
End If
        Me.Caption = "Now Reading - " + .filetitle
        SB.Panels(1) = .filename
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
        RTF.Tag = .filename
        Undo.Reset 'clear undo buffer
        FileChanged = False 'reset need to save flag
        RTF.SelStart = 0
        NoStatusUpdate = False
        EditEnable
        LockWindowUpdate 0
        Screen.MousePointer = 0 'hourglass
    End With
      
End Sub
Private Sub mnuFilePageSetup_Click()
    ShowPageSetupDlg
End Sub
Private Sub mnuFilePrint_Click()
On Error Resume Next
  ShowPrinter
End Sub
Private Sub mnuFileProperties_Click()
    Dim Temp As Variant, z As Long, count As Long, Msg As String
    Dim charcnt As Long, linecnt As Long, Mcount As Long
    If RTF.Tag <> "" And FileExists(RTF.Tag) Then
        'do a windows Property Dialog if we can
        GetPropDlg Me, RTF.Tag
    Else
        'otherwise just give document statistics
        linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'line count
        charcnt = Len(RTF.Text) 'character count
        Temp = Split(RTF.Text, Chr(32)) 'word count
        For z = 0 To UBound(Temp)
            Select Case Trim(Temp(z))
                Case vbNullString
                Case vbCrLf
                Case vbCr
                Case Else
                    Mcount = Mcount + 1
            End Select
        Next z
        Msg = "File not yet saved." + vbCrLf
        Msg = Msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
        Msg = Msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
        Msg = Msg + "Lines :" + Format(linecnt, "#,###,###,##0")
        MsgBox Msg, vbInformation, "Michael's NotePad"
    End If
End Sub
Private Sub mnuFileSave_Click()
    SaveAFile
End Sub
Private Sub mnuFileSaveAs_Click()
   Dim sFile As String
   ' Dim Response As VbMsgBoxResult
   
        ' Don't do anything if a document isn't open
        If Me Is Nothing Then Exit Sub
        '
           ' Save file as
        'With CDl
                  RTF.Tag = sFile
                CDl.dialogtitle = "Save The Document As ?"
               ' .CancelError = False
                'ToDo: set the flags and attributes of the common dialog control
                CDl.Filter = "Plain Text Document (*.txt)|*.txt|RTF Document (*.rtf)|*.rtf|Professional NotePad Document (*.pnt)|*.pnt|All Files (*.*)|*.*"
                '.ShowSave
                CDl.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
    CDl.CancelError = False
    CDl.ShowSave
   
                If Len(CDl.filename) = 0 Then
                        Exit Sub
                End If
                 sFile = CDl.filename
        'Me.Caption = sFile
        'Me.RTF.SaveFile sFile
        'If FileChanged =  Then
        'If Not SaveAFile Then
        'Me.RTF.SaveFile sFile
        RTF.SaveFile sFile
         'End If
        FileChanged = False 'reset flag
        Me.Caption = "Now Reading - " + FileOnly(sFile)
        SB.Panels(1) = sFile
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
    ' Create a "Save File" dialog box
     If Err Then Exit Sub
   
    'End With
End Sub
Private Sub mnuFormatBackcolor_Click()
    Dim col As Long 'new backcolor
    col = ShowColor
    If col <> -1 Then
        If col < 1 Then col = -col
        RTF.BackColor = col
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Backcolor", Str(col)
    End If
End Sub
Private Sub mnuFormatFont_Click()
    On Error GoTo woops
    Dim st As Long, curvl As Long, FontChange As Boolean
    'FileChanged is set to true by RTF change event (in the class)
    'As we are only changing font - not content, we dont want
    'this to alter due to this sub, so remember current state
    'so we can reset it below
    ChangeState = FileChanged
    Undo.IgnoreChange True  'dont add this action to the Undo buffer
    'current position
    curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    st = RTF.SelStart
    With SelectFont
        .mFontName = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Fontname", "Lucida Console")
        .mFontsize = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "FontSize", "9"))
        .mBold = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Bold", False))
        .mItalic = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Italic", False))
        .mStrikethru = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "StrikeThru", False))
        .mUnderline = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Underline", False))
        .mFontColor = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Color", Str(vbBlack)))
        ShowFont
        If .mFontName <> GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Fontname", "Lucida Console") Then FontChange = True
        If .mFontsize <> .mFontsize = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "FontSize", "9")) Then FontChange = True
        If .mBold <> .mBold = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Bold", False)) Then FontChange = True
        If .mItalic <> .mItalic = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Italic", False)) Then FontChange = True
        If .mStrikethru <> .mStrikethru = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "StrikeThru", False)) Then FontChange = True
        If .mUnderline <> .mUnderline = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Underline", False)) Then FontChange = True
        If .mFontColor <> .mFontColor = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Color", Str(vbBlack))) Then FontChange = True
        If Not FontChange Then GoTo woops
        'save new font
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Fontname", .mFontName
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "FontSize", Str(.mFontsize)
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Bold", .mBold
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Italic", .mItalic
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "StrikeThru", .mStrikethru
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Underline", .mUnderline
        SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Color", Str(.mFontColor)
        'implement on our new font
        LockWindowUpdate Me.hwnd
        RTF.SelStart = 0
        RTF.SelLength = Len(RTF.Text)
        RTF.SelColor = .mFontColor
        RTF.SelFontName = .mFontName
        RTF.SelFontSize = .mFontsize
        RTF.SelBold = .mBold
        RTF.SelItalic = .mItalic
        RTF.SelStrikeThru = .mStrikethru
        RTF.SelUnderline = .mUnderline
        RTF.SelStart = st
        RTF.SelLength = 0
        SetScrollPos curvl, RTF 'reset to current scroll position
    End With
woops:
    Undo.IgnoreChange False  'start using Undo system again
    FileChanged = ChangeState 'reset to what it was
    RTF.SetFocus
    LockWindowUpdate 0
End Sub

Private Sub mnuFormatStats_Click()
    Dim Temp As Variant, z As Long, count As Long, Msg As String
    Dim charcnt As Long, linecnt As Long, Mcount As Long
    linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'lines
    charcnt = Len(RTF.Text) 'characters
    Temp = Split(RTF.Text, Chr(32)) 'words
    For z = 0 To UBound(Temp)
        Select Case Trim(Temp(z))
            Case vbNullString
            Case vbCrLf
            Case vbCr
            Case Else
                Mcount = Mcount + 1
        End Select
    Next z
    Msg = IIf(RTF.Tag = "", "File not yet saved.", RTF.Tag) + vbCrLf
    Msg = Msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
    Msg = Msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
    Msg = Msg + "Lines :" + Format(linecnt, "#,###,###,##0")
    MsgBox Msg, vbInformation, "Michael's NotePad"
End Sub

Private Sub mnuFormatWordwrap_Click()
    mnuFormatWordwrap.Checked = Not mnuFormatWordwrap.Checked
    RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
End Sub

Private Sub mnuPlay2_Click()
'mnuPlay2EffectEffect.Checked = Not mnuPlay2EffectEffect.Checked
  If mnuPlay2.Checked = True Then
mnuPlay2.Checked = False

Else
mnuPlay2.Checked = True
 gHW = Me.hwnd
    gsMidiFile = App.Path & "\Music\Music.mid"
    gbRepeat = True  'Set to true if you want it to repeat
    
    'You can specify the start and stop times with
    'these variables.  To Play the whole song, just
    'set both glFrom and glTo to 0.
    glFrom = 0
    glTo = 0 'Will play for a couple of seconds
    
    'Go ahead and make sure no other midi file is open
    CloseMidi
    UnHook
    DoEvents
    
    'It is important to hook the form after
    'you begin playing the Midi.  Not sure why.
    PlayMidi
    Hook
    End If
If mnuPlay2.Checked = False Then
CloseMidi
UnHook
mnuPlay2.Caption = "Not Listening To Soothing Music - (Click Here to Enable Play2)..."
Else
mnuPlay2.Caption = "Listen to Soothing Music While You Type... (Click To Toggle!))"
End If
End Sub

Private Sub mnuPlay_Click()
If mnuPlay.Checked = True Then
mnuPlay.Checked = False
Else
mnuPlay.Checked = True
End If
If mnuPlay.Checked = False Then
mnuPlay.Caption = "Not Using Sound Effects! - (Click Here to Enable Sound Effects)..."
Else
mnuPlay.Caption = "Using Sound Effects! - (Click Here To Disable Sound Effects!)"
End If
End Sub

Private Sub mnuPopupAbout_Click()
mnuFileAbout_Click
End Sub

Private Sub mnuRestore_Click()
 Dim Result As Long
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
End Sub

Private Sub mnuSkin_Click()
CDl.Filter = "Mike's NotePad Themes (*.uis, *.sss)|*.uis;*.sss|All files|*.*"
CDl.dialogtitle = "Load A Cool Theme For Mike's NotePad!"
CDl.initdir = App.Path = "\Skins\"
CDl.Flags = CDl.Flags Or cdlOFNPathMustExist Or cdlOFNFileMustExist
CDl.ShowOpen
If Not Len(CDl.filename) = 0 Then
    Wb.SetRootPath 1
    Dim first As Integer, second As Integer, cur As Integer
    cur = 3
    Do While cur > 0
        first = second
        second = cur
        cur = InStr(cur + 1, CDl.filename, "\")
    Loop
    Wb.SetRootPathStr Left$(CDl.filename, first)
    m_SkinName = Right$(CDl.filename, Len(CDl.filename) - first)

    LoadSkin m_SkinName = Right$(CDl.filename, Len(CDl.filename) - first)
    'cmbSkin.ListIndex = -1
    ' SaveSetting "Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Skin", m_SkinName
End If
End Sub
Private Sub LoadSkin(skinname As String)
    ' Give DirectSkin the relative path of the new skin to load
    Wb.LoadUIS m_SkinName
    ' Reload
    Wb.ReloadUIS
    If MsgBox("Would You like to Set This Theme As Your Default Theme?", vbQuestion Or vbYesNo, "Set As Default?") = vbYes Then
    
    SaveSetting "MichaelsKoolNotePadXp\" + App.EXEName, "Settings", "Choose A Different Theme...", m_SkinName
   
    MsgBox "The Theme You Have Selected Is Now Your Default Theme...", vbInformation, "The Theme Has Been Set!"
     'Else
     
     End If

    'Wb.SetWindowTheme (hwnd),
    'Call SaveSetting(App.EXEName, "Skins", "Current", filename)
End Sub


Private Sub mnuSplashScreen_Click()
If mnuSplashScreen.Checked = True Then
mnuSplashScreen.Checked = False
Else
mnuSplashScreen.Checked = True
End If
If mnuSplashScreen.Checked = False Then
mnuSplashScreen.Caption = "Splash Screen Is Disabled... - (Click Here to Enable)..."
Else
mnuSplashScreen.Caption = "Splash Screen Is Enabled... - (Click Here To Disable The Splash Screen)"
End If
End Sub

Private Sub mnuSystray_Click()
 If mnuSystray.Checked = True Then
mnuSystray.Checked = False

Else
mnuSystray.Checked = True
'SetWindowLong Me.hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, _
 '       GWL_EXSTYLE) And Not WS_EX_APPWINDOW)

 With nid
  
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Michael J. Hardy's Skinnable Windows NotePad!" & vbNullChar
       
       Shell_NotifyIcon NIM_ADD, nid
       End With
    End If
If mnuSystray.Checked = False Then
  SetWindowLong Me.hwnd, GWL_EXSTYLE, (GetWindowLong(hwnd, _
        GWL_EXSTYLE) Or WS_EX_APPWINDOW)

 Shell_NotifyIcon NIM_DELETE, nid

mnuSystray.Caption = "Minimize To System Tray..."
Else
mnuSystray.Caption = "Minimize To System Tray..."
End If
End Sub

Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    SB.Visible = mnuViewStatusbar.Checked
End Sub
Private Sub mnuViewToolbar_Click()
 '  mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
 ' mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked = True Then
mnuViewToolbar.Checked = False
Else
mnuViewToolbar.Checked = True
End If
 
If mnuViewToolbar.Checked = False Then
mnuViewToolbar.Caption = "Toolbar is Invisible... - (Click Here to Enable The Toolbar!)..."
Else
mnuViewToolbar.Caption = "The Toolbar is Visible... - (Click Here To Close the Toolbar!)"
End If
    tbrMain.Visible = mnuViewToolbar.Checked
End Sub

Private Sub PasteTimer_Timer()
    'you could hook the clipboard, but this will do
    mnuEdit(5).Enabled = Clipboard.GetFormat(vbCFText)
    tbrMain.Buttons(8).Enabled = mnuEdit(5).Enabled
End Sub
Private Sub PicLeft_Resize()
    On Error Resume Next
    RTF.Width = PicLeft.Width - 120
    RTF.Height = PicLeft.Height
End Sub

Private Sub RTF_Change()
  
   If NoStatusUpdate Then Exit Sub
  
  
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength

End Sub

Private Sub RTF_GotFocus()
    Dim z As Long 'allow tabs WITHIN richtextbox
    ReDim mTStop(0 To Controls.count - 1) As Boolean
    On Local Error Resume Next
    If mnuPlay.Checked = True Then
PlaySound 110
End If
    For z = 0 To Controls.count - 1
        mTStop(z) = Controls(z).TabStop
        Controls(z).TabStop = False
    Next
    SelFont
End Sub

Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)
    SelFont
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
    If mnuPlay.Checked = True Then
PlaySound Typer
Else
End If
End Sub

Private Sub RTF_LostFocus()
    Dim z As Long 'reset tabstops to original state
    On Local Error Resume Next
    For z = 0 To Controls.count - 1
        Controls(z).TabStop = mTStop(z)
    Next
End Sub

Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     'If mnuPlay.Checked = True Then
'PlaySound 109

    SelFont
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If mnuPlay.Checked = True Then
PlaySound 106
    If Button = 2 Then Me.PopupMenu mnuEditBase
End If
'Me.PopupMenu mnuEditBase

End Sub

Private Sub RTF_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Response As VbMsgBoxResult, Temp As String
    If Data.GetFormat(vbCFFiles) Then
        If FileChanged Then 'do we save current doc ?
         If mnuPlay.Checked = True Then
PlaySound 107
End If
            Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
           
            Select Case Response
                Case vbCancel
                    Effect = vbDropEffectNone
                    Exit Sub
                Case vbYes
                
                    If Not SaveAFile Then
                        Effect = vbDropEffectNone
                        Exit Sub
                    End If
            End Select
        End If
        'Data.Files is a collection of the filepaths of files
        'dropped onto a control. Multiple files may be dropped
        'but in this app, we can only open one at a time
        'so we are only interested in Data.Files(1)
        Temp = Data.Files(1)
        Temp = strUnQuoteString(Temp) 'sometimes explorer uses quotes('send to' for example)
        Temp = GetLongFilename(Temp) 'looks better than a dos path
        RTF.Text = ""
        SelFont
        Select Case LCase(ExtOnly(Temp))
            Case "txt"
                RTF.SelText = OneGulp(Temp) 'binary read
            Case "rtf"
                RTFtemp.LoadFile Temp 'rtf load
                RTF.SelText = RTFtemp.Text
            Case "doc"
                OpenWordDoc Temp 'see sub - returns plain text
            Case Else
                RTF.SelText = OneGulp(Temp) 'otherwise do binary read
        End Select
        Me.Caption = FileOnly(Temp)
        SB.Panels(1) = Temp
        SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
        RTF.Tag = Temp
        Undo.Reset
        FileChanged = False
    Else
        Effect = vbDropEffectNone
    End If
     If mnuPlay.Checked = True Then
PlaySound 108
End If
End Sub

Private Sub RTF_SelChange()
    If NoStatusUpdate Then Exit Sub
    EditEnable
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    If mnuPlay.Checked = True Then
PlaySound 105
End If
    'see menu items for comments
    Select Case Button.Index
    
        Case 1
            mnuFileNew_Click
        Case 2
            mnuFileOpen_Click
        Case 3
            mnuFileSave_Click
        Case 4
            mnuFileSaveAs_Click
        Case 6
            mnuEdit_Click 3
        Case 7
            mnuEdit_Click 4
        Case 8
            mnuEdit_Click 5
        Case 9
            mnuEdit_Click 6
        Case 11
            mnuEdit_Click 8
        Case 13
            mnuEdit_Click 0
        Case 14
            mnuEdit_Click 1
        Case 16
            mnuEdit_Click 13
        Case 18
            mnuEdit_Click 15
            Case 19
            mnuFilePrint_Click
            Case 20
            mnuFileAbout_Click
    End Select
End Sub

Private Sub TB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    mnuEdit_Click ButtonMenu.Index + 7
End Sub

Private Sub TrayIcon_OnMenuItemSelect(ByVal ItemIndex As Long)
'Text1.Text = Text1.Text + CStr(ItemIndex) + vbNewLine
    With frmMain
    If ItemIndex = 1010000000 Then
        mnuFileAbout_Click
    End If
    If ItemIndex = 1100000000 Then
    mnuFileExit_Click
        'TrayIcon.DetachFromSysTray
       ' End
    End If
        End With
    'If ItemIndex = 1090000000 Then
     '   TrayIcon.DetachFromSysTray
      '  Unload Me
       ' End
    'End If
    

End Sub



Private Sub tbrMain_ButtonClick(ByVal Button As ComctlLib.Button)
 If mnuPlay.Checked = True Then
PlaySound 105
End If
    'see menu items for comments
    Select Case Button.Index
    
        Case 1
            mnuFileNew_Click
        Case 2
            mnuFileOpen_Click
        Case 3
            mnuFileSave_Click
        Case 4
            mnuFileSaveAs_Click
        Case 6
            mnuEdit_Click 3
        Case 7
            mnuEdit_Click 4
        Case 8
            mnuEdit_Click 5
        Case 9
            mnuEdit_Click 6
        Case 11
            mnuEdit_Click 8
        Case 13
            mnuEdit_Click 0
           ' mnuEdit_Click 15
        Case 14
            mnuEdit_Click 1
        Case 16
            mnuEdit_Click 13
        Case 17
            mnuEdit_Click 15
            Case 18
            mnuFilePrint_Click
            Case 20
            mnuSkin_Click
            Case 22
            mnuFileAbout_Click
    End Select
End Sub

Private Sub Undo_StateChanged()
    'enable menus/toolbar buttons according to undo class
    mnuEdit(0).Enabled = Undo.canUndo
    mnuEdit(1).Enabled = Undo.canRedo
    tbrMain.Buttons(13).Enabled = Undo.canUndo
    tbrMain.Buttons(14).Enabled = Undo.canRedo
    'Set FileChanged flag so we know if we need to save
    FileChanged = (Undo.canUndo = True Or Undo.canRedo = True)
End Sub
Private Sub OpenWordDoc(mfile As String)
    'standard call to Word to open a file
    'and get just the text
    Dim WordApp As Object
    On Error GoTo woops
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open mfile
    WordApp.ActiveDocument.Content.Copy
    SelFont
    RTF.SelText = Clipboard.GetText(vbCFText)
    WordApp.Application.Quit
    Set WordApp = Nothing
    Screen.MousePointer = 0
    Exit Sub
woops:
    Set WordApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Private Sub SaveAsWordDoc(mfile As String)
    ' get Word to save as .doc file
    'Why bother ? Well firstly, this is a demo of functionality
    'and I thought some people might like to know how to do this,
    'and secondly, when the file gets opened by Word in the future
    'sometimes Word is not happy with the fact that a text file
    'has a .doc extension and prompts for a new 'filter'
    'to be installed from CD or some other complaint - but still
    'opens the file. So... save as Word Document, avoid problems.
    Dim WordApp As Object
    Dim Document As Object
    On Error GoTo woops
    Screen.MousePointer = 11
    Set WordApp = CreateObject("Word.Application")
    Set Document = WordApp.Documents.Add
    Clipboard.Clear
    Clipboard.SetText RTF.Text, vbCFText
    WordApp.ActiveDocument.Content.Paste
    Document.SaveAs mfile
    WordApp.Application.Quit
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    Exit Sub
woops:
    Set WordApp = Nothing
    Set Document = Nothing
    Screen.MousePointer = 0
    MsgBox "Error converting Word Document", vbCritical
End Sub
Public Function SaveAFile() As Boolean
    Dim Response As VbMsgBoxResult, sFile As String
    If Not FileExists(RTF.Tag) Then
       ' GoTo DoSaveAs 'must be a new file
         mnuFileSaveAs_Click
    Else
        Select Case LCase(ExtOnly(RTF.Tag))
            Case "rtf"
                'if it was an existing .rtf file it will lose
                'formatting because we're only plain text
                'even though we'll save in Rich text format
                Response = MsgBox("Any rich text formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        SaveAFile = False
                        Exit Function
                    Case vbYes
                       mnuFileSaveAs_Click ' GoTo DoSaveAs
                End Select
                RTF.SaveFile RTF.Tag
            Case "doc"
                'same as .rtf
                Response = MsgBox("Any document formatting in this file will be lost." + vbCrLf + "Do you wish to save this file using a different name ?", vbYesNoCancel)
                Select Case Response
                    Case vbCancel
                        SaveAFile = False
                        Exit Function
                    Case vbYes
                      mnuFileSaveAs_Click
                        'GoTo DoSaveAs
                End Select
                SaveAsWordDoc RTF.Tag
            Case Else 'just plain text
                Kill RTF.Tag
                FileSave RTF.Text, RTF.Tag
        End Select
        FileChanged = False
        SaveAFile = True
    End If
    Exit Function
DoSaveAs:
    With cmndlg
        .filefilter = "Plain text (*.txt)|*.txt|Rich text (*.rtf)|*.rtf|Word Document (*.doc)|*.doc|All files (*.*)|*.*"
        .Flags = 5 Or 2
        SaveFile
        If Len(.filename) = 0 Then
            SaveAFile = False
            Exit Function
        End If
        'make sure we have the correct extension
        Select Case .filefilterindex
            Case 1
                If InStr(1, sFile, ".") = 0 Then
                    sFile = sFile + ".txt"
                Else
                    sFile = ChangeExt(sFile, "txt")
                End If
                FileSave RTF.Text, sFile 'plain text
            Case 2
                If InStr(1, sFile, ".") = 0 Then
                    sFile = sFile + ".rtf"
                Else
                    sFile = ChangeExt(sFile, "rtf")
                End If
                RTF.SaveFile sFile 'rich text format
            Case 3
                If InStr(1, sFile, ".") = 0 Then
                    sFile = sFile + ".doc"
                Else
                    sFile = ChangeExt(sFile, "doc")
                End If
                SaveAsWordDoc sFile 'word document
            Case 4
                If InStr(1, sFile, ".") = 0 Then sFile = sFile + ".txt"
                FileSave RTF.Text, .filename 'plain text
        End Select
        Me.Caption = .filetitle
        SB.Panels(1) = .filename
        SB.Panels(4) = GetFileSize(Len(RTF.Text))
        RTF.Tag = .filename
        FileChanged = False 'reset flag
    End With
    SaveAFile = True
    
End Function

Public Sub SelFont()
    'This a plain text editor - just this font thanks
    'It is far more efficient to call this fast routine often
    'than do a SelectAll, ChangeFont, Select none routine -
    'particularly with large files
    RTF.SelFontName = GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Fontname", "Lucida Console")
    RTF.SelFontSize = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "FontSize", "9"))
    RTF.SelBold = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Bold", False))
    RTF.SelItalic = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Italic", False))
    RTF.SelStrikeThru = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "StrikeThru", False))
    RTF.SelUnderline = CBool(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Underline", False))
    RTF.SelColor = Val(GetSetting("Michael-J-Hardy's-Professional-Software\" + App.EXEName, "Settings", "Color", Str(vbBlack)))
End Sub


Public Sub EditEnable()
    'enable menus/toolbar buttons according to selection length
    Dim Enabled As Boolean
    Enabled = (RTF.SelLength > 0)
    mnuEdit(3).Enabled = Enabled
    mnuEdit(4).Enabled = Enabled
    mnuEdit(6).Enabled = Enabled
    tbrMain.Buttons(6).Enabled = Enabled
    tbrMain.Buttons(7).Enabled = Enabled
    tbrMain.Buttons(9).Enabled = Enabled
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub
