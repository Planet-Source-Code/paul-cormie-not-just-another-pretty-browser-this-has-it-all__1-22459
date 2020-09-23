VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmPad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IB v1.10-Text Box"
   ClientHeight    =   7335
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12165
   Icon            =   "FrmPad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12165
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2640
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   7200
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save As"
      Filter          =   "Text Files (*.txt)|*.txt"
      InitDir         =   "C:\"
   End
   Begin MSComctlLib.ImageList imlToolbarIcons2 
      Left            =   5520
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":158C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":169E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":17B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":1ACA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":1BDC
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":1CEE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":1E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":211A
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":222C
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":233E
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":2450
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":276A
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":287C
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":298E
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":2AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":2DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":30D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":33EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":3708
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":3A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":3D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4056
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4168
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4482
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4594
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":46A6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":47B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4AD2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4BE4
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4CF6
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":4E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5122
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5234
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5346
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5458
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5772
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":5DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPad.frx":60C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   6720
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   11880
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"FrmPad.frx":63DA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7065
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11086
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "Internet Browser 1.1"
            TextSave        =   "Internet Browser 1.1"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   952
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   952
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   952
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   952
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            TextSave        =   "1/16/01"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons2"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create a new file."
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open existing file."
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save current file"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print current file."
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo / Redo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo last action and if pressed again, will Redo that action"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select All"
            Description     =   "Select All"
            Object.ToolTipText     =   "Select Everything in Document"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut selected text."
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy to clipboard."
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste from clipboard."
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete selected text."
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold font."
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic font."
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline font."
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find text in document."
            Object.ToolTipText     =   "Find"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Description     =   "Align text left."
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Align text center."
            Object.ToolTipText     =   "Center"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Description     =   "Align text right."
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Color"
            Object.ToolTipText     =   "Font Color"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Time"
            Description     =   "Time"
            Object.ToolTipText     =   "Click for current time"
            ImageIndex      =   17
            Style           =   1
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Source"
            Description     =   "View Source"
            Object.ToolTipText     =   "View Source"
            ImageIndex      =   20
         EndProperty
      EndProperty
      Begin VB.ComboBox size 
         Height          =   315
         ItemData        =   "FrmPad.frx":64C7
         Left            =   10920
         List            =   "FrmPad.frx":64FE
         TabIndex        =   10
         Text            =   "10"
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox text 
         Height          =   315
         ItemData        =   "FrmPad.frx":6545
         Left            =   8640
         List            =   "FrmPad.frx":6576
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Verdana"
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSComctlLib.Toolbar tbFind 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   1111
      ButtonWidth     =   1482
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      MousePointer    =   14
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1440
         Picture         =   "FrmPad.frx":6635
         ScaleHeight     =   495
         ScaleMode       =   0  'User
         ScaleWidth      =   132
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   8520
         Picture         =   "FrmPad.frx":693F
         ScaleHeight     =   495
         ScaleMode       =   0  'User
         ScaleWidth      =   132
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "Search For:"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "&Next"
         Height          =   330
         Left            =   7200
         TabIndex        =   6
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "F&ind"
         Height          =   330
         Left            =   6240
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   3210
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Back to IB"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditTimeDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSerpator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFonts 
         Caption         =   "F&onts"
         Begin VB.Menu mnuFontsBold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnuFontsItalic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnuFontsUnderline 
            Caption         =   "&Underline"
         End
         Begin VB.Menu mnuFontsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFontsFont 
            Caption         =   "&Font..."
         End
         Begin VB.Menu mnuFontsColor 
            Caption         =   "&Color..."
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewViewSource 
         Caption         =   "&View HTML Source from Browser"
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Cedits"
      End
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'
' Developed by Paul Cormie
' paul_cormie @ hotmail.com
'
' Canadian, eh
'
'====================================================================================
'
' *****  READ THIS BEFORE USING THIS CODE:  ******
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Internet Browser fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author.
'
' The source code for Internet Browser version 1.10 has been submitted
' for the purposes of education.  I find the best way to learn is to
' look at how other people do things and see if i can possibly do it
' more efficiently. Contact me for additional help/suggestions via my
' website. There is a form to contact me there.
'
' VISIT MY WEBSITE : http://www.paul_cormie.homestead.com
'
'===================================================================================='
'
'********************************************
'********************************************
'******    DECLAIRATIONS START HERE    ******
'********************************************
'********************************************

Option Explicit
Dim DocChanged As Boolean
Dim docname As String

'********************************************
'********************************************
'******    DECLAIRATIONS ENDS HERE     ******
'********************************************
'********************************************










'********************************************
'********************************************
'****     FORM CONTROLS STARTS HERE    ******
'********************************************
'********************************************

Private Sub Form_Activate()
    ChangeToolBar                                                ' Updates toolbar and menus.
    ChangeMenus
End Sub

Private Sub Form_Load()
    Dim LineWidth As Long
        ChangeToolBar
        ChangeMenus                                              ' Updates toolbar and menus.
            docname = " (Untitled)"
            Me.Caption = "Text Box " & docname
        cmdFind.Enabled = False                                  ' Disables the Find and FindNext buttons
        cmdFindNext.Enabled = False
        RichTextBox1.SelFontName = text.text
        RichTextBox1.SelFontSize = CSng(size.text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DocChanged Then                                           ' Checks to see if document has changed since last save
        Select Case MsgBox("The file has changed." & vbCr & vbCr & "Do you wish to save your changes?", vbExclamation + vbYesNoCancel, frmPad.Caption) ' If it has gives a message box with a chance to save
        Case vbYes
            mnuFileSave_Click
        Case vbNo
            Unload frmPad
        Case vbCancel
            Cancel = True
        End Select
    End If
End Sub


Private Sub cmdFind_Click()
Dim textfound As Integer
    cmdFindNext.Enabled = True                                   ' Enables the cmdFindNext
   
    RichTextBox1.Find (Text1.text)                               ' Finds the text in the search box and highlights it,
    RichTextBox1.SetFocus                                        ' Sets the focus on the richtextbox so the selected text is editable
    textfound = RichTextBox1.Find(Text1.text)                    ' The richtextbox1.find method returns an integer value of -1 if the searched for text is not found.
        If textfound = -1 Then
            MsgBox "End of Document" & vbCr & vbCr & "Text Not Found", vbInformation, App.Title  ' If this is true then it displays a message box.
        End If
End Sub

Private Sub cmdFindNext_Click()
    mnuEditFindNext_Click
End Sub

Private Sub RichTextBox1_Change()
    DocChanged = True                                            ' Changes the docchanged value to true for saving purposes.
    ChangeToolBar                                                ' Updates the toolbar and menus.
    ChangeMenus
End Sub

Private Sub RichTextBox1_Click()
    ChangeToolBar                                                ' Updates the toolbar and menus just in case.
    ChangeMenus
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeMenus                                                  ' Updates the toolbar and menus just in case.
    ChangeToolBar
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ChangeMenus                                                  ' Updates the toolbar and menus just in case.
    ChangeToolBar
End Sub

Private Sub CopytoClipBoard()
    Clipboard.SetText RichTextBox1.SelText                        ' Copies selected text to clipboard.
End Sub

Public Sub DeleteSelectedText()
    RichTextBox1.SelText = ""                                     ' Sets the value of the selected text to nothing.
End Sub

Private Sub size_Click()
    RichTextBox1.SelFontSize = CSng(size.text)
End Sub

Private Sub text_Click()
    RichTextBox1.SelFontName = text.text
End Sub

Private Sub Text1_Change()
    cmdFind.Enabled = True                                        ' Enables the find button when the search text is entered.
End Sub

'********************************************
'********************************************
'****     FORM CONTROLS ENDS HERE      ******
'********************************************
'********************************************










'********************************************
'********************************************
'***     TOOLBAR  STARTS  HERE       ********
'********************************************
'********************************************

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
        Select Case Button.Key                                   ' Just tells the toolbar buttons what to do...
            
            Case "New"
                mnuFileNew_Click
            Case "Open"
                mnuFileOpen_Click
            Case "Save"
                mnuFileSave_Click
            Case "Print"
                mnuFilePrint_Click
            Case "Undo / Redo"
                SendKeys "^z"
            Case "Select All"
                    RichTextBox1.SelStart = 0                    ' Selects all the text in the current document.
                    RichTextBox1.SelLength = Len(RichTextBox1.text)
                    ChangeToolBar                                ' Updates the toolbar and menus.
                    ChangeMenus
            Case "Cut"
                mnuEditCut_Click
            Case "Copy"
                mnuEditCopy_Click
            Case "Paste"
                mnuEditPaste_Click
            Case "Delete"
                mnuEditDelete_Click
            Case "Bold"
                mnuFontsBold_Click
            Case "Italic"
                mnuFontsItalic_Click
            Case "Underline"
                mnuFontsUnderline_Click
            Case "Find"
                mnuEditFind_Click
            Case "Align Left"
                RichTextBox1.SelAlignment = rtfLeft              ' Aligns text to left margin.
            Case "Center"
                RichTextBox1.SelAlignment = rtfCenter            ' Aligns text to center.
            Case "Align Right"
                RichTextBox1.SelAlignment = rtfRight             ' Aligns text to right margin.
            Case "Font Color"
                    CDL1.Flags = cdlCCFullOpen                   ' Shows the color dialogue box and sets the current color.
                    CDL1.ShowColor
                    RichTextBox1.SelColor = CDL1.Color
            Case "Time"                                          ' FrmTime opens and stays on top of all other forms
                    If FrmTime.Visible = True Then
                        tbToolBar.Buttons(27).Value = tbrUnpressed
                        FrmTime.Visible = False
                    Else
                        tbToolBar.Buttons(27).Value = tbrPressed
                        FrmTime.Visible = True
                    End If
            Case "View Source"
                Dim src As String
                RichTextBox1.text = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text) ' FrmInternetBrowser.cboAddress.text =URL address input
                src = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text)         ' stores source to the varible src
    End Select
End Sub

Public Sub ChangeToolBar()
    If RichTextBox1.SelBold Then                                  ' Makes portions of the tool bar context sensative.
        tbToolBar.Buttons("Bold").Value = tbrPressed
    Else
        tbToolBar.Buttons("Bold").Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelItalic Then
        tbToolBar.Buttons("Italic").Value = tbrPressed
    Else
        tbToolBar.Buttons("Italic").Value = tbrUnpressed
    End If
    
    If RichTextBox1.SelUnderline Then
        tbToolBar.Buttons("Underline").Value = tbrPressed
    Else
        tbToolBar.Buttons("Underline").Value = tbrUnpressed
    End If
End Sub

'********************************************
'********************************************
'***     TOOLBAR  ENDS  HERE         ********
'********************************************
'********************************************










'********************************************
'********************************************
'****       MENU  STARTS  HERE     **********
'********************************************
'********************************************

Private Sub mnuHelpAbout_Click()
    FrmCredits.Show
End Sub

Private Sub mnuEditDelete_Click()
    DeleteSelectedText                                             'Updates toolbar and menus.
    ChangeToolBar
    ChangeMenus
End Sub

Private Sub mnuEditFind_Click()
    If tbFind.Visible = False Then
        tbFind.Visible = True                                      ' Shows the Find Toolbar.
            If tbToolBar.Visible = False Then
                RichTextBox1.Move 0, 630, Me.ScaleWidth, Me.ScaleHeight - 900
            Else
                RichTextBox1.Move 0, 1050, Me.ScaleWidth, Me.ScaleHeight - 1320
            End If
    Else
        tbFind.Visible = False
            If tbToolBar.Visible = False Then                     ' Hides the Find Toolbar.
                RichTextBox1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 270
            Else
                RichTextBox1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690
            End If
    End If
End Sub

Private Sub mnuEditFindNext_Click()
    RichTextBox1.SetFocus                                         ' Set the focus so the selected text can be directly edited.
    RichTextBox1.Find (Text1.text), RichTextBox1.SelStart + 1
End Sub

Private Sub mnuEditSelectAll_Click()
    RichTextBox1.SelStart = 0                                     ' Selects all the text in the current document.
    RichTextBox1.SelLength = Len(RichTextBox1.text)
    ChangeToolBar
    ChangeMenus
End Sub

Private Sub mnuEditTimeDate_Click()
    Dim text As String
    Dim SelStart As Long

    DeleteSelectedText                                           ' Deletes the selected text, if any, and gets ready to insert the time and date string.
    
    If RichTextBox1.SelLength > 0 Then
    End If
           
    text = RichTextBox1.text                                     ' Inserts time and date.
    SelStart = RichTextBox1.SelStart
    RichTextBox1.text = Left(text, SelStart) & Now & Right(text, Len(text) - SelStart)
    RichTextBox1.SelStart = SelStart                             ' Resets cursor to original position.
    ChangeToolBar                                                ' Updates toolbar and menus.
    ChangeMenus
End Sub

Private Sub mnuFontsBold_Click()
    If RichTextBox1.SelBold Then                                 ' Sets the Bold property and makes the menu context sensative.
        RichTextBox1.SelBold = False
        mnuFontsBold.Checked = False
    Else
        RichTextBox1.SelBold = True
        mnuFontsBold.Checked = True
    End If
End Sub

Private Sub mnuFontsColor_Click()
    CDL1.Flags = cdlCCFullOpen                                  ' Shows the color dialog box and sets the current color.
    CDL1.ShowColor
    RichTextBox1.SelColor = CDL1.Color
End Sub

Private Sub mnuFontsFont_Click()
    CDL1.Flags = cdlCFBoth Or cdlCFEffects                      ' Shows the Font dialog box and sets the current font.
    CDL1.ShowFont
    With RichTextBox1
        .SelFontName = CDL1.FontName
        .SelFontSize = CDL1.FontSize
        .SelBold = CDL1.FontBold
        .SelItalic = CDL1.FontItalic
        .SelStrikeThru = CDL1.FontStrikethru
        .SelUnderline = CDL1.FontUnderline
        .SelColor = CDL1.Color
    End With
End Sub

Private Sub mnuFontsItalic_Click()
    If RichTextBox1.SelItalic Then                               ' Sets the italic property and makes the menu context sensative.
        RichTextBox1.SelItalic = False
        mnuFontsItalic.Checked = False
    Else
        RichTextBox1.SelItalic = True
        mnuFontsItalic.Checked = True
    End If
End Sub

Private Sub mnuFontsUnderline_Click()
    If RichTextBox1.SelUnderline Then                            ' Sets the underline property and makes the menu context sensative.
        RichTextBox1.SelUnderline = False
        mnuFontsUnderline.Checked = False
    Else
        RichTextBox1.SelUnderline = True
        mnuFontsUnderline.Checked = True
    End If
End Sub

Private Sub mnuViewViewSource_Click()
    Dim src As String
    RichTextBox1.text = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text) ' FrmInternetBrowser.cboAddress.text =URL address input
    src = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text)     ' stores source to the varible src
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked    ' Shows or hides the status bar as needed.
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked        ' Shows or hides the toolbar as needed.
    tbToolBar.Visible = mnuViewToolbar.Checked
       
    If tbToolBar.Visible = False Then                          ' This resizes the richtextbox depending on the state of the toolbar.
        RichTextBox1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 270
    Else
        RichTextBox1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 690
    End If
End Sub

Private Sub mnuEditPaste_Click()
        
    Dim text As String
    Dim ClipboardText As String
    Dim SelStart As Long
        
    If Clipboard.GetFormat(vbCFText) Then
        If RichTextBox1.SelLength > 0 Then                     ' Replace selected text. (if any)
            DeleteSelectedText
        End If
        text = RichTextBox1.text                               ' Move text we need to a variable.
        SelStart = RichTextBox1.SelStart
        ClipboardText = Clipboard.GetText
        
        RichTextBox1.text = Left(text, SelStart) & ClipboardText & Right(text, Len(text) - SelStart)
        RichTextBox1.SelStart = SelStart                       ' Restore the cursor position.
    Else
        ChangeMenus
        ChangeToolBar
    End If
End Sub

Private Sub mnuEditCopy_Click()
    CopytoClipBoard
    ChangeMenus
    ChangeToolBar
End Sub

Private Sub mnuEditCut_Click()
    CopytoClipBoard
    DeleteSelectedText
    ChangeMenus
    ChangeToolBar
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    Dim bcancel As Boolean
    Dim ncopy As Integer
    On Error GoTo errorhandler
    
    bcancel = False
    
    CDL1.Flags = cdlPDHidePrintToFile Or _
            cdlPDNoSelection Or cdlPDNoPageNums _
            Or cdlPDCollate
    
    CDL1.CancelError = True
    CDL1.PrinterDefault = True
    CDL1.Copies = 1
    CDL1.ShowPrinter
    
        If bcancel = False Then
            PrintRTF RichTextBox1, 1440, 1440, 1440, 1440
            For ncopy = 1 To CDL1.Copies
            Next ncopy
        End If
    
    Exit Sub
    
errorhandler:
    If Err.Number = cdlCancel Then
    bcancel = True
    Resume Next
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim Cancel As Boolean
    On Error GoTo errorhandler
    Cancel = False
    
    CDL1.DefaultExt = ".txt"
    CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
    CDL1.CancelError = True
    CDL1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    
    CDL1.ShowSave
    
    If Not Cancel Then
        If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
            RichTextBox1.SaveFile CDL1.FileName, rtfRTF
        Else
            RichTextBox1.SaveFile CDL1.FileName, rtfText
        End If
        RichTextBox1.FileName = CDL1.FileName
        docname = CDL1.FileName
        Me.Caption = App.Title & " " & docname
        DocChanged = False
    End If
    
    Exit Sub
    
errorhandler:
    If Err.Number = cdlCancel Then
        Cancel = True
        Resume Next
    End If
End Sub

Private Sub mnuFileSave_Click()
    If docname = " (Untitled)" Then
        mnuFileSaveAs_Click
    Else
        If UCase(Right(RichTextBox1.FileName, 3)) = "RTF" Then
            RichTextBox1.SaveFile RichTextBox1.FileName, rtfRTF
        Else
            RichTextBox1.SaveFile RichTextBox1.FileName, rtfText
        End If
        DocChanged = False
    End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim Cancel As Boolean
    On Error GoTo errorhandler
    Cancel = False
    
    CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files|*.*"
    CDL1.CancelError = True
    CDL1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    CDL1.ShowOpen
    
    If Not Cancel Then
        If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
            RichTextBox1.LoadFile CDL1.FileName, rtfRTF
        Else
            RichTextBox1.LoadFile CDL1.FileName, rtfText
        End If
            RichTextBox1.FileName = CDL1.FileName
            docname = RichTextBox1.FileName
            Me.Caption = App.Title & " " & docname
            DocChanged = False
    End If
    Exit Sub
    
errorhandler:
    If Err.Number = cdlCancel Then
        Cancel = True
        Resume Next
    End If
    End
End Sub

Private Sub mnuFileNew_Click()
Dim Cancel As Integer

If DocChanged = False Then
    RichTextBox1.text = ""
Else
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, frmPad.Caption)
    
    Case vbYes
        mnuFileSave_Click
    Case vbNo
        RichTextBox1.text = ""
    Case vbCancel
        Cancel = True
    
    End Select
End If

End Sub

Public Sub ChangeMenus()

    mnuFileSave.Enabled = DocChanged                              ' Makes the menus and toolbar context sensative.
    mnuFileSaveAs.Enabled = DocChanged
    mnuEditCopy.Enabled = False
    mnuEditCut.Enabled = False
    mnuEditDelete.Enabled = False
    mnuEditPaste.Enabled = False
    tbToolBar.Buttons("Save").Enabled = DocChanged
        
    If RichTextBox1.text = "" Then                                'If RichTextBox1.SelLength > 0 Then
        mnuEditCut.Enabled = False
        mnuEditCopy.Enabled = False
        mnuEditDelete.Enabled = False
        tbToolBar.Buttons("Cut").Enabled = False
        tbToolBar.Buttons("Copy").Enabled = False
        tbToolBar.Buttons("Delete").Enabled = False
        '
        tbToolBar.Buttons("Save").Enabled = False
        tbToolBar.Buttons("Select All").Enabled = False
        tbToolBar.Buttons("Undo / Redo").Enabled = False
        tbToolBar.Buttons("Print").Enabled = False
        tbToolBar.Buttons("Find").Enabled = False
        
    Else
        mnuEditCut.Enabled = True
        mnuEditCopy.Enabled = True
        mnuEditDelete.Enabled = True
        tbToolBar.Buttons("Cut").Enabled = True
        tbToolBar.Buttons("Copy").Enabled = True
        tbToolBar.Buttons("Delete").Enabled = True
        '
        tbToolBar.Buttons("Save").Enabled = True
        tbToolBar.Buttons("Select All").Enabled = True
        tbToolBar.Buttons("Undo / Redo").Enabled = True
        tbToolBar.Buttons("Print").Enabled = True
        tbToolBar.Buttons("Find").Enabled = True
        
    End If
    
    If Clipboard.GetFormat(vbCFText) Then
        mnuEditPaste.Enabled = True
        tbToolBar.Buttons("Paste").Enabled = True
    Else
        mnuEditPaste.Enabled = False
        tbToolBar.Buttons("Paste").Enabled = True
    End If
    
    If RichTextBox1.SelBold Then
        mnuFontsBold.Checked = True
    Else
        mnuFontsBold.Checked = False
    End If
    
    If RichTextBox1.SelItalic Then
        mnuFontsItalic.Checked = True
    Else
        mnuFontsItalic.Checked = False
    End If
    
    If RichTextBox1.SelUnderline Then
        mnuFontsUnderline.Checked = True
    Else
        mnuFontsUnderline.Checked = False
    End If
    
End Sub


'********************************************
'********************************************
'****       MENU  ENDS  HERE     ************
'********************************************
'********************************************







'-------------------------------------------------
'   temp storage area  ---  START
'-------------------------------------------------




'-------------------------------------------------
'   temp storage area  ---  END
'-------------------------------------------------


