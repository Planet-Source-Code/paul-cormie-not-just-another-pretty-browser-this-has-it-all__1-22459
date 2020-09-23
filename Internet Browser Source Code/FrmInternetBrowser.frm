VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmInternetBrowser 
   Caption         =   "Internet Browser Version 1.10"
   ClientHeight    =   8250
   ClientLeft      =   3210
   ClientTop       =   3900
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInternetBrowser.frx":0000
   LinkTopic       =   "FrmInternetBrowser"
   ScaleHeight     =   8250
   ScaleWidth      =   12960
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10440
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer timerAnimate 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   12000
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":18A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":1BC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   11520
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   10215
      Begin VB.CommandButton cmdTextBox 
         Caption         =   "&Advanced Features"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         ToolTipText     =   "Opens a more advanced text editor called ""Text Box"""
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdViewSource 
         Caption         =   "View Source"
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save &As"
         Height          =   372
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Save your document..."
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   372
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdCut 
         Caption         =   "Cut"
         Height          =   372
         Left            =   2280
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   372
         Left            =   1320
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtQuick 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1480
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   0
         Width           =   6615
      End
      Begin VB.Shape Shape1 
         Height          =   1440
         Left            =   50
         Top             =   10
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   720
         Picture         =   "FrmInternetBrowser.frx":1EDA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmInternetBrowser.frx":314C
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11040
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7995
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   450
      SimpleText      =   "Format(Now, ""long time"")"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12038
            MinWidth        =   1058
            Text            =   "Done"
            TextSave        =   "Done"
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
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   952
            MinWidth        =   952
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "3/8/01"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.TreeView TreeHistory 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4683
      _Version        =   393217
      LabelEdit       =   1
      Style           =   1
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   795
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   1402
      BandCount       =   4
      FixedOrder      =   -1  'True
      _CBWidth        =   12975
      _CBHeight       =   795
      _Version        =   "6.0.8450"
      Child1          =   "Toolbar1"
      MinHeight1      =   390
      Width1          =   9675
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Address"
      Child2          =   "cboAddress"
      MinHeight2      =   315
      Width2          =   7995
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Caption3        =   "VB Links"
      Child3          =   "cboURLList"
      MinHeight3      =   315
      FixedBackground3=   0   'False
      NewRow3         =   0   'False
      AllowVertical3  =   0   'False
      Caption4        =   "JavaScript Links"
      Child4          =   "cboJS"
      MinHeight4      =   315
      FixedBackground4=   0   'False
      NewRow4         =   0   'False
      AllowVertical4  =   0   'False
      Begin VB.ComboBox cboJS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmInternetBrowser.frx":31ED
         Left            =   10275
         List            =   "FrmInternetBrowser.frx":3221
         TabIndex        =   17
         Text            =   "Select a JavaScript Site"
         Top             =   450
         Width           =   2610
      End
      Begin VB.ComboBox cboAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Text            =   "Enter URL"
         Top             =   450
         Width           =   7170
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   688
         ButtonWidth     =   2223
         ButtonHeight    =   688
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList2"
         HotImageList    =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Back"
               Key             =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Forward"
               Key             =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Stop"
               Key             =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Home"
               Key             =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Text Box"
               Key             =   "TextBox"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "History"
               Key             =   "History"
               ImageIndex      =   7
               Style           =   1
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search"
               Key             =   "Search"
               ImageIndex      =   8
               Style           =   1
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Bookmarks"
               Key             =   "Bookmarks"
               ImageIndex      =   9
               Style           =   1
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboURLList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmInternetBrowser.frx":3489
         Left            =   8880
         List            =   "FrmInternetBrowser.frx":34D2
         TabIndex        =   6
         Text            =   "Select a VB Site"
         ToolTipText     =   "Select the dropdown box for quick shortcuts"
         Top             =   450
         Width           =   0
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   10440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save As"
      Filter          =   "Text Files (*.txt)|*.txt"
      InitDir         =   "C:\"
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11160
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":37E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":3D3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":4299
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":47F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":4D51
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":52AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":55C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":5B97
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":60F3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   11880
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":640D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":6969
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":6EC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":7421
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":797D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":7ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":81F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":86BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInternetBrowser.frx":8C17
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSeperator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenNewWindow 
         Caption         =   "&Open New Window               "
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCommands 
         Caption         =   "&Commands"
         Begin VB.Menu mnuFileCommandsStop 
            Caption         =   "&Stop"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuFileCommandsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileCommandsBack 
            Caption         =   "&Back"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuFileCommandsForward 
            Caption         =   "&Forward               "
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuFileCommandsSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileCommandsHome 
            Caption         =   "&Home"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuFileCommandsRefresh 
            Caption         =   "&Refresh"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu mnuFileSeperator8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFilePrintPrint 
            Caption         =   "Pint..."
         End
         Begin VB.Menu mnuFilePrintPageSetUp 
            Caption         =   "Page Setup..."
         End
      End
      Begin VB.Menu mnuFileSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTextBox 
         Caption         =   "&Text Box               "
         Begin VB.Menu mnuViewTextBoxShowSimple 
            Caption         =   "&Show / Hide Simple Text Box               "
            Shortcut        =   {F12}
         End
         Begin VB.Menu mnuViewSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewTextBoxShowComplex 
            Caption         =   "&Show Complex Text Box               "
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuViewTextBoxHideComplex 
            Caption         =   "&Hide Complex Text Box"
         End
      End
      Begin VB.Menu mnuViewSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "&History"
         Begin VB.Menu mnuViewHistoryShowHistory 
            Caption         =   "&Show / Hide History               "
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuViewSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "&Search"
         Begin VB.Menu mnuViewSearchShowSearch 
            Caption         =   "&Show Search               "
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuViewSearchHideSearch 
            Caption         =   "&Hide Search"
            Shortcut        =   +{F8}
         End
      End
      Begin VB.Menu mnuViewSeperator4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewMP3Player 
         Caption         =   "MP3 Player"
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewSeperator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDisablePopUps 
         Caption         =   "Disable other windows from poping up"
      End
   End
   Begin VB.Menu mnuInternetTools 
      Caption         =   "&Internet Tools"
      Begin VB.Menu mnuInternetToolsListAllTools 
         Caption         =   "&List All Tools"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuInternetToolsOpenAllTools 
         Caption         =   "&Open All Tools"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuInternetToolsCloseAllTools 
         Caption         =   "&Close All Tools"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuInternetToolsSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternetToolsPing 
         Caption         =   "&Ping"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuInternetToolsPortScanner 
         Caption         =   "Port &Scanner"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuInternetToolsPortListener 
         Caption         =   "Port &Listener"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuInternetToolsHostLookup 
         Caption         =   "&Host Lookup"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuInternetToolsTraceRoute 
         Caption         =   "&Trace Route"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuInternetToolsWhois 
         Caption         =   "&Whois"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuInternetToolsViewHTMLSourceCode 
         Caption         =   "&View HTML Source Code"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuViewBookmarksShowBookmarks 
         Caption         =   "&Show Bookmarks               "
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewBookmarksHideBookmarks 
         Caption         =   "&Hide  Bookmarks"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuViewSeperator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBookmarksClearBookmarks 
         Caption         =   "&Clear Bookmarks"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHints 
         Caption         =   "&Hints"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuHelpSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpBugs 
         Caption         =   "Bugs?..."
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisclaimer 
         Caption         =   "&Disclaimer               "
         Begin VB.Menu mnuDisclaimerReadme 
            Caption         =   "&Readme"
         End
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "FrmInternetBrowser"
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
'****    DECLAIRATIONS STARTS HERE     ******
'********************************************
'********************************************

Option Explicit
Dim strHome As String
Dim dayAdded As Boolean
Dim NewLocation As String
Dim Today As String
Dim TodayInHistory As Integer
Dim ThisDayName As String
Dim SlashNumber
Dim Position
Dim OldLocation As String
Dim KeyNumber
Dim DayNumber As Integer
Dim nodCN As Node
Dim nodUrl As Node
Dim Length
Public AllowPopups As Boolean


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

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        CoolBar1.Width = FrmInternetBrowser.Width
        ProgressBar1.Width = FrmInternetBrowser.Width
        
    On Error Resume Next
        If TreeHistory.Visible = True Then
            TreeHistory.Height = FrmInternetBrowser.Height - 1350
            WebBrowser1.Width = FrmInternetBrowser.Width - 3290
            WebBrowser1.Height = FrmInternetBrowser.Height - 1350
            WebBrowser1.Left = 3290
            ProgressBar1.Width = FrmInternetBrowser.Width - 100
            ProgressBar1.Left = 50
        Else
            If Frame1.Visible = True Then
                Frame1.Height = FrmInternetBrowser.Height - 1350
                WebBrowser1.Width = FrmInternetBrowser.Width - 3290
                WebBrowser1.Height = FrmInternetBrowser.Height - 1350
                WebBrowser1.Left = 3290
            Else
                TreeHistory.Height = FrmInternetBrowser.Height - 1350
                WebBrowser1.Width = FrmInternetBrowser.Width - 100
                WebBrowser1.Height = FrmInternetBrowser.Height - 1350
                WebBrowser1.Left = 50
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    
    TodayInHistory = 0                                          ' TodayInHistory will be used in the LoadHistory sub
    
    FrmInternetBrowser.WindowState = vbMaximized                ' Maximize the window
    TreeHistory.Visible = False
    ProgressBar1.Width = Me.ScaleWidth                          ' Sizes the progress bar
    Frame1.Visible = False                                      ' dayAdded will be used in the AddToday sub
    dayAdded = False
    TreeHistory.Nodes.Clear
    TestToday                                                   ' Call the TestToday sub
    LoadHistory                                                 ' Call the LoadHistory sub
    DeleteHistory                                               ' Call the DeleteHistory sub
    AddToday                                                    ' Call the AddToday sub
    WebBrowser1.GoHome
    AllowPopups = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
' Save the history when exiting by calling the
' SaveHistory sub
SaveHistory

End Sub

Private Sub Flash1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command                                         ' Loads the whole size of the FLASH movie into the buffer,not just the animation
    End Select
End Sub

Private Sub Web_StatusTextChange(ByVal text As String)
    StatusBar1.Panels(1).text = text                            ' Shows the page being loaded.
End Sub




'********************************************
'********************************************
'****     FORM CONTROLS ENDS HERE    ********
'********************************************
'********************************************










'********************************************
'********************************************
'****       MENU  STARTS HERE    ************
'********************************************
'********************************************
Private Sub mnuFileSave_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub mnuDisclaimerReadme_Click()
    MsgBox "All trademarks appearing on Internet Browser 1.1 are trademarks of their respective owners.", 8, ""
End Sub

Private Sub mnuFilePrintPageSetUp_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub mnuFileCommandsBack_Click()
    On Error Resume Next
    WebBrowser1.GoBack
End Sub

Private Sub mnuFilePrintPrint_Click()
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuFileCommandsForward_Click()
    On Error Resume Next
    WebBrowser1.GoForward
End Sub

Private Sub mnuFileCommandsHome_Click()
    WebBrowser1.GoHome
End Sub

Private Sub mnuFileCommandsRefresh_Click()
    On Error Resume Next
    WebBrowser1.Refresh
End Sub

Private Sub mnuFileCommandsStop_Click()
    On Error Resume Next
    WebBrowser1.Stop
End Sub

Private Sub mnuFileExit_Click()
    End                                                         ' This extremely high end VB command is unknown as to what function it preforms
End Sub

Private Sub mnuFileOpenNewWindow_Click()
    
    On Error Resume Next
    Static lDocumentCount As Long
    Dim IB2 As Form
    lDocumentCount = lDocumentCount + 1
    Set IB2 = New FrmInternetBrowser
    IB2.Show
    IB2.SetFocus
    
End Sub

Private Sub mnuHelpAbout_Click()
    FrmCredits.Show
End Sub

Private Sub mnuHelpBugs_Click()
    FrmBugs.Show
End Sub

Private Sub mnuHelpHelp_Click()
    Dim strMsgBox As String                                     ' Message Box prompts that help is not available, then asks if you'd like to go to
    Dim strMsgBox2 As String                                    ' my site for tech help.  If you say OK, it takes you to my site
    strMsgBox = MsgBox("Help not available in Version 1.1" & vbCrLf & vbCrLf & "The next release will have a help file included." & vbCrLf & vbCrLf & "Contact me from my website if you have questions.", 8, "")
        If strMsgBox = "2" Then
        Else
            strMsgBox2 = MsgBox("Would you like to go there now?", vbOKCancel, "")
            If strMsgBox2 = "2" Then
            Else
                cboAddress.text = "http://www.paul_cormie.homestead.com"
                WebBrowser1.Navigate cboAddress.text
            End If
        End If
End Sub

Private Sub mnuViewTextBoxHideSimple_Click()
        ' See the toolbar Text Box button for details on what this does
        
        If Frame1.Visible = True Then
            Toolbar1.Buttons(9).Value = tbrUnpressed
            Frame1.Visible = False
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 1280
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 1200
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 960
                TreeHistory.Top = 960
            End If
               
        Else
            Toolbar1.Buttons(9).Value = tbrPressed
            Frame1.Visible = True
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 2900
            WebBrowser1.Left = 0
            WebBrowser1.Top = 2600
            Frame1.Width = Me.ScaleWidth
            txtQuick.Width = Me.ScaleWidth - 3400
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 2600
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 2600
                TreeHistory.Top = 2600
                TreeHistory.Left = 0
            End If
        End If
        
End Sub

Private Sub mnuHelpHints_Click()
        FrmHint.Show
End Sub

Private Sub mnuInternetToolsHostLookup_Click()
    FrmLookup.Show
End Sub

Private Sub mnuInternetToolsListAllTools_Click()
    FrmInternetTools.Show
End Sub

Private Sub mnuInternetToolsOpenAllTools_Click()
    FrmPing.Show
    FrmPing.Show
    FrmListen.Show
    FrmScanner.Show
    FrmTrace.Show
    FrmGetHTML.Show
    FrmWhois.Show
End Sub

Private Sub mnuInternetToolsCloseAllTools_Click()
    FrmPing.Hide
    FrmPing.Hide
    FrmListen.Hide
    FrmScanner.Hide
    FrmTrace.Hide
    FrmGetHTML.Hide
    FrmWhois.Hide
End Sub

Private Sub mnuInternetToolsPing_Click()
    FrmPing.Show
End Sub

Private Sub mnuInternetToolsPortListener_Click()
    FrmListen.Show
End Sub

Private Sub mnuInternetToolsPortScanner_Click()
    FrmScanner.Show
End Sub

Private Sub mnuInternetToolsTraceRoute_Click()
    FrmTrace.Show
End Sub

Private Sub mnuInternetToolsViewHTMLSourceCode_Click()
    FrmGetHTML.Show
End Sub

Private Sub mnuInternetToolsWhois_Click()
    FrmWhois.Show
End Sub

Private Sub mnuViewBookmarksClearBookmarks_Click()
Dim MsgBoxClear As String
    
    MsgBoxClear = MsgBox("Are you sure you want to clear the bookmarks?", vbOKCancel, "Clear ?")
    
    If MsgBoxClear = 2 Then
      MsgBox "Bookmarks *NOT* Cleared", 8, ""
    Else
        On Error Resume Next
        FrmBookmarks.lstBookmarks.Clear
        Dim i As Integer
        Dim a As String
        Open App.Path & "\bookmarks.txt" For Output As #1
        For i = 0 To FrmBookmarks.lstBookmarks.ListCount - 1
        Write #1, ""
        Next i
        Close #1
        FrmBookmarks.lstBookmarks.text = "http://www.paul_cormie.homestead.com"
        MsgBox "Bookmarks Cleared", 8, ""
    End If
End Sub

Private Sub mnuViewBookmarksHideBookmarks_Click()
        Toolbar1.Buttons(15).Value = tbrUnpressed
        FrmBookmarks.Visible = False
End Sub

Private Sub mnuViewBookmarksShowBookmarks_Click()
        Toolbar1.Buttons(15).Value = tbrPressed
        FrmBookmarks.Visible = True
End Sub

Private Sub mnuViewDisablePopUps_Click()
    If mnuViewDisablePopUps.Checked = False Then
        mnuViewDisablePopUps.Checked = True
        AllowPopups = False
    ElseIf mnuViewDisablePopUps.Checked = True Then
        mnuViewDisablePopUps.Checked = False
        AllowPopups = True
    End If
End Sub

Private Sub mnuViewHistoryShowHistory_Click()
        ' See the toolbar history button for details on what this does
        
        If TreeHistory.Visible = True Then
            Toolbar1.Buttons(11).Value = tbrUnpressed
            TreeHistory.Visible = False
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 1200
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
            
            If Frame1.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth
                WebBrowser1.Height = Me.ScaleHeight - 3000
                WebBrowser1.Left = 0
                WebBrowser1.Top = 2600
                TreeHistory.Top = 960
                TreeHistory.Height = Me.ScaleHeight
            End If
               
        Else
            Toolbar1.Buttons(11).Value = tbrPressed
            TreeHistory.Visible = True
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
                    
            If Frame1.Visible = False Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 1200
                WebBrowser1.Left = 3290
                WebBrowser1.Top = 960
                TreeHistory.Top = 960
                TreeHistory.Height = Me.ScaleHeight
                TreeHistory.Left = 0
            Else
                If Frame1.Visible = True Then
                    WebBrowser1.Width = Me.ScaleWidth - 3290
                    WebBrowser1.Height = Me.ScaleHeight - 2900
                    WebBrowser1.Left = 3290
                    WebBrowser1.Top = 2600
                    TreeHistory.Top = 2600
                    TreeHistory.Left = 0
                    
                End If
            End If
        End If
End Sub

Private Sub mnuViewMP3Player_Click()
    MsgBox "I have decided not to put a useless MP3 player in this version of Internet Browser." & vbCrLf & vbCrLf & "I felt there was no point in putting an average player in this project when you can use a quality one like Winamp for free" & vbCrLf & "Next version, if popular eunough, will include a quality media player in in.", 8, ""
    'frmPlayer.Show
    
End Sub

Private Sub mnuViewSearchShowSearch_Click()
        Toolbar1.Buttons(13).Value = tbrPressed
        FrmSearch.Visible = True
End Sub

Private Sub mnuViewSearchHideSearch_Click()
        Toolbar1.Buttons(13).Value = tbrUnpressed
        FrmSearch.Visible = False
End Sub

Private Sub mnuViewTextBoxHideComplex_Click()
    If frmPad.Visible = True Then
        frmPad.Hide
    Else
    End If
End Sub

Private Sub mnuViewTextBoxShowComplex_Click()
    frmPad.Show
End Sub

Private Sub mnuViewTextBoxShowSimple_Click()
        ' Same code exactly as the show text box
        ' All it is is an if then else a few times
        
        If Frame1.Visible = True Then
            Toolbar1.Buttons(9).Value = tbrUnpressed
            Frame1.Visible = False
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 1280
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 1200
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 960
                TreeHistory.Top = 960
            End If
               
        Else
            Toolbar1.Buttons(9).Value = tbrPressed
            Frame1.Visible = True
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 2900
            WebBrowser1.Left = 0
            WebBrowser1.Top = 2600
            Frame1.Width = Me.ScaleWidth
            txtQuick.Width = Me.ScaleWidth - 3400
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 2600
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 2600
                TreeHistory.Top = 2600
                TreeHistory.Left = 0
            End If
        End If
        
End Sub

'********************************************
'********************************************
'****       MENU  ENDS  HERE     ************
'********************************************
'********************************************










'********************************************
'********************************************
'***     TOOLBAR  STARTS  HERE       ********
'********************************************
'********************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

' All the different buttons and commands for each one
Select Case Button
    Case "Back"
        On Error Resume Next
        WebBrowser1.GoBack
    Case "Forward"
        On Error Resume Next
        WebBrowser1.GoForward
    Case "Stop"
        WebBrowser1.Stop
    Case "Refresh"
        WebBrowser1.Refresh
    Case "Home"
        WebBrowser1.GoHome
    Case "Search"
        If FrmSearch.Visible = True Then
            Toolbar1.Buttons(13).Value = tbrUnpressed
            FrmSearch.Visible = False
        Else
            Toolbar1.Buttons(13).Value = tbrPressed
            FrmSearch.Visible = True
        End If
        
    Case "History"
    ' The following code is used to show (or hide) the
    ' History, resize the WebBrowser, the textbox and the
    ' History, and to change the value of the
    ' History button to Pressed or Unpressed depending
    ' on the TreeHistory visibility.....similar code for the
    ' TextBox button
    '
    ' This one was bloody tricky to get the both working right!
    
        If TreeHistory.Visible = True Then
            Toolbar1.Buttons(11).Value = tbrUnpressed
            TreeHistory.Visible = False
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 1200
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
            
            If Frame1.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth
                WebBrowser1.Height = Me.ScaleHeight - 3000
                WebBrowser1.Left = 0
                WebBrowser1.Top = 2600
                TreeHistory.Top = 960
                TreeHistory.Height = Me.ScaleHeight
            End If
               
        Else
            Toolbar1.Buttons(11).Value = tbrPressed
            TreeHistory.Visible = True
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
                    
            If Frame1.Visible = False Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 1200
                WebBrowser1.Left = 3290
                WebBrowser1.Top = 960
                TreeHistory.Top = 960
                TreeHistory.Height = Me.ScaleHeight
                TreeHistory.Left = 0
            Else
                If Frame1.Visible = True Then
                    WebBrowser1.Width = Me.ScaleWidth - 3290
                    WebBrowser1.Height = Me.ScaleHeight - 2900
                    WebBrowser1.Left = 3290
                    WebBrowser1.Top = 2600
                    TreeHistory.Top = 2600
                    TreeHistory.Left = 0
                    
                End If
            End If
        End If
    Case "Text Box"
        'See the history button for details on what this does
        
        If Frame1.Visible = True Then
            Toolbar1.Buttons(9).Value = tbrUnpressed
            Frame1.Visible = False
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 1280
            WebBrowser1.Left = 0
            WebBrowser1.Top = 960
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 1200
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 960
                TreeHistory.Top = 960
            End If
               
        Else
            Toolbar1.Buttons(9).Value = tbrPressed
            Frame1.Visible = True
            WebBrowser1.Width = Me.ScaleWidth
            WebBrowser1.Height = Me.ScaleHeight - 2900
            WebBrowser1.Left = 0
            WebBrowser1.Top = 2600
            Frame1.Width = Me.ScaleWidth
            txtQuick.Width = Me.ScaleWidth - 3400
            
            If TreeHistory.Visible = True Then
                WebBrowser1.Width = Me.ScaleWidth - 3290
                WebBrowser1.Height = Me.ScaleHeight - 2600
                WebBrowser1.Left = 3240
                WebBrowser1.Top = 2600
                TreeHistory.Top = 2600
                TreeHistory.Left = 0
            End If
        End If
        
        Case "Bookmarks"
       
        If FrmBookmarks.Visible = True Then
            Toolbar1.Buttons(15).Value = tbrUnpressed
            FrmBookmarks.Visible = False
                        
        Else
            Toolbar1.Buttons(15).Value = tbrPressed
            FrmBookmarks.Visible = True
        End If
        
    End Select
End Sub


Private Sub cboJS_KeyPress(KeyAscii As Integer)
' All three of the coolbar listboxes use the same idea.
' When the enter button (KeyAscii = 13) is pressed
' the value in either three of the boxes is used
' OldLocation will be used in the urlTest sub
' Call the urlTest sub --> urlTest

' JavaScript List
    If KeyAscii = 13 Then                                           ' Press Enter will execute the URL
        OldLocation = cboJS.text
        cboJS.AddItem cboAddress.text
        WebBrowser1.Navigate cboJS.text
        urlTest
    End If
End Sub

Private Sub cboURLList_KeyPress(KeyAscii As Integer)

' VB List
    If KeyAscii = 13 Then                                           ' Press Enter will execute the URL
        OldLocation = cboURLList.text
        cboURLList.AddItem cboAddress.text
        WebBrowser1.Navigate cboURLList.text
        urlTest
    End If
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
'URL List
    If KeyAscii = 13 Then                                           ' Press Enter will execute the URL
        OldLocation = cboAddress.text
        cboAddress.AddItem cboAddress.text
        If OldLocation <> "" Then
        WebBrowser1.Navigate cboAddress.text
        End If
        urlTest
    End If
End Sub


'********************************************
'********************************************
'***     TOOLBAR  ENDS   HERE        ********
'********************************************
'********************************************










'********************************************
'********************************************
'***  WEBBROWSER CONTROLS STARTS HERE   *****
'********************************************
'********************************************

Private Sub WebBrowser1_StatusTextChange(ByVal text As String)
    If text <> "Done" Then
        StatusBar1.Panels(1).text = text                    ' When the site is finished loading the status bar says "done"
    End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, url As Variant)
Dim Text1
    FrmInternetBrowser.Caption = WebBrowser1.LocationName & "-" & WebBrowser1.LocationURL
    Text1 = WebBrowser1.LocationURL                         ' Shows the url in text1
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error Resume Next
    ProgressBar1.Max = ProgressMax                          ' Progress bar above the browser
    ProgressBar1.Value = Progress
    ProgressBar1.Refresh
End Sub

'********************************************
'********************************************
'***  WEBBROWSER CONTROLS ENDS HERE     *****
'********************************************
'********************************************










'********************************************
'********************************************
'****  TEXTBOX CONTROLS STARTS  HERE   ******
'********************************************
'********************************************

Private Sub cmdViewSource_Click()
    Dim src As String
    If FrmInternetBrowser.cboAddress.text = "about:blank" Then
        MsgBox "The page is blank", 8, ""
    Else
        txtQuick.text = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text)   ' FrmInternetBrowser.cboAddress.text =URL address input
        src = Inet1.OpenURL(FrmInternetBrowser.cboAddress.text)             ' stores source to the varible src
    End If
End Sub

Private Sub cmdTextBox_Click()
    frmPad.Show
End Sub

Private Sub Image1_Click()
    FrmCredits.Show
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear                                         ' Delete everthing in the clipboard
    Clipboard.SetText txtQuick.SelText, 1                   ' Put your text into it on place 1
End Sub

Private Sub cmdCut_Click()
    Clipboard.Clear                                         ' Delete everthing in the clipboard
    Clipboard.SetText txtQuick.SelText, 1                   ' Put your text into it on place 1
                                                            ' If the place 1 allready excists it will be erased
    txtQuick.SelText = ""                                   ' Delete everyting that was selected in the textbox
End Sub

Private Sub cmdPaste_Click()
    txtQuick.SelText = Clipboard.GetText(1)                 ' Get the text in the clipboard on place 1 and
                                                            ' Place it on the selected area in the textbox
                                                            ' If nothing is selected it will be place on the place of writing cursor
End Sub

Private Sub cmdSave_Click()                                 ' Common dialog box opens and prompts to save
    On Error Resume Next
    dlgSave.ShowSave
        If dlgSave.FileName = "" Then                       ' If no name then it won't save
        Else
            Call SaveText(txtQuick, dlgSave.FileName)       ' Otherwise call Savetext function
        End If
End Sub

Sub SaveText(Lst As TextBox, File As String)
Dim X
    On Error GoTo error                                      ' Call SaveText (txtQuick,"C:\________temp_file.txt")
    Dim mystr As String
    Open File For Output As #1
    Print #1, Lst
    Close 1
    MsgBox "The File:" & vbCrLf & vbCrLf & dlgSave.FileName & vbCrLf & vbCrLf & "has been saved", 8, "Confirmation"
    Exit Sub
error:                                                      ' Error handling
    X = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub

Private Sub cmdClear_Click()
Dim strMsgBox
Dim i
Dim txtWordCount

    If txtQuick.text = "" Then                                ' Confirmation of clearing the text box
        MsgBox " Already Clear ", 8, ""
    Else
           strMsgBox = MsgBox(" Are you sure you want to clear the text box?", vbYesNo, "")
           If strMsgBox = vbNo Then
           Else
               For i = 1 To Me.Controls.Count - 1
                   If TypeOf Me.Controls(i) Is TextBox Then
                   Me.Controls(i).text = ""
                   txtWordCount = ""
                   End If
                   Next i
           End If
    End If
End Sub

'********************************************
'********************************************
'****  TEXTBOX CONTROLS ENDS  HERE     ******
'********************************************
'********************************************










'********************************************
'********************************************
'***     HISTORY  STARTS   HERE      ********
'********************************************
'********************************************

' written by Hani Karam
' see credits page for contact info
Public Sub OneSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located before the first
' slash if there is no "http://" in it
While SlashNumber = 0
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 1, the computer name is
    ' found
    If (SlashNumber = 0) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Public Sub TwoSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com"
    ' If the slash number is 2, add "/" at the end of
    ' the URL so it can be used in the
    ' ThreeSlashURL sub because if the slash number
    ' is smaller than 3, we will have an infinite loop
    OldLocation = OldLocation & "/"
    ' Call the ThreeSlashURL sub
    ThreeSlashURL

End Sub

Public Sub ThreeSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located between the
' "http://" and the next slash, which makes the slash
' number equals to 3
While SlashNumber < 3
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 2, the computer name
    ' begins
    If (SlashNumber = 2) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Private Sub SaveHistory()
Open App.Path & "\history.txt" For Output As #1
Dim currentNode As Node
For Each currentNode In TreeHistory.Nodes
Select Case currentNode.text
    Case "Sunday"
        Print #1, "Sunday"
    Case "Monday"
        Print #1, "Monday"
    Case "Tuesday"
        Print #1, "Tuesday"
    Case "Wednesday"
        Print #1, "Wednesday"
    Case "Thursday"
        Print #1, "Thursday"
    Case "Friday"
        Print #1, "Friday"
    Case "Saturday"
        Print #1, "Saturday"
    Case Else
        ' If currentNode.text is not a day, it might
        ' either a computer name or a complete URL
        If currentNode.Children > 0 Then
        ' currentNode.Children > 0 means currentNode.text
        ' is a computer name, then print one tab
            Print #1, vbTab; currentNode.text
        Else
        ' currentNode.text is a complete URL, then print
        ' two tabs
            Print #1, vbTab; vbTab; currentNode.text
        End If
End Select
Next currentNode
Close #1
End Sub

Private Sub LoadHistory()
Dim fnum
Dim file_name
Dim text_line
Dim level
Dim num_nodes


' This code will search a text file for tab to create a
' treeview depending on the tab number.
' I found this code on PSC, but I modified it to suit my
' needs
Dim tree_nodes() As Node
fnum = FreeFile
' Initialize KeyNumber and DayNumber
KeyNumber = 1
DayNumber = 1
' Open the history file
file_name = App.Path & "\history.txt"
    Open file_name For Input As fnum
    
    TreeHistory.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        Select Case level
        Case 1
        ' If Level = 1, that means we have a day name
            Set tree_nodes(level) = TreeHistory.Nodes.Add(, , "day" & KeyNumber, text_line, 1)
                ' keyNumber will be used later in this
                ' sub and in the DeleteHistory sub
                KeyNumber = KeyNumber + 1
                If text_line = Today Then
                    ' Expand the day node
                    tree_nodes(level).Expanded = True
                    ' TodayInHistory will be used later
                    ' in this sub and in the
                    ' DeleteHistory sub
                    TodayInHistory = 1
                    ' dayAdded will be used in the
                    ' AddToday sub
                    dayAdded = True
                    ' Today will be used in the AddToday
                    ' sub
                    Today = "day" & (KeyNumber - 1)
                    ' DayNumber will be used in the
                    ' DeleteHistory sub
                    DayNumber = KeyNumber
                End If
        Case 2
        ' If Level = 2, that means we have a computer
        ' name
            ' If TodayInHistory = 0, that means that
            ' today's name was not added to the history
            ' tree from the saved file yet. For example,
            ' if today is Wednesday and the loaded node
            ' is Tuesday (or Monday), this node will not
            ' be used to add a URL to history while
            ' using the browser, it will be used only
            ' to view the saved history, that's why there
            ' is no need to create a key for it
            If TodayInHistory = 0 Then
                Set tree_nodes(level) = TreeHistory.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line, 2, 3)
            Else
            ' If TodayInHistory = 1, that means that today's
            ' name was added from the saved file, so a key is
            ' necessary to prevent adding a URL that is already
            ' in the history tree
                Set tree_nodes(level) = TreeHistory.Nodes.Add(tree_nodes(level - 1), tvwChild, text_line, text_line, 2, 3)
            End If
        Case Else
            ' Same explanation as above
            If TodayInHistory = 0 Then
                Set tree_nodes(level) = TreeHistory.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line, 4)
            Else
                Set tree_nodes(level) = TreeHistory.Nodes.Add(tree_nodes(level - 1), tvwChild, text_line, text_line, 4)
            End If
       End Select
    Loop

    Close fnum
End Sub


Public Sub TestToday()
' This sub is used to find out what day is today
Select Case Weekday(Now())
    Case 1
        Today = "Sunday"
    Case 2
        Today = "Monday"
    Case 3
        Today = "Tuesday"
    Case 4
        Today = "Wednesday"
    Case 5
        Today = "Thursday"
    Case 6
        Today = "Friday"
    Case 7
        Today = "Saturday"
End Select
' ThisDayName will be used in the AddToday sub
ThisDayName = Today
End Sub

Public Sub urlTest()
SlashNumber = 0
NewLocation = ""

Length = Len(OldLocation)
' Count the slash number
For Position = 1 To Length
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
Next Position

Select Case SlashNumber
    Case 0
    ' Example : www.yahoo.com
        ' If there are not any slashes in the URL then
        ' there is no need to change it
        NewLocation = OldLocation
        ' If a slash is not added at the end of
        ' OldLocation, this will generate en error as
        ' NewLocation and OldLocation are used as keys
        ' in the TreeHistory
        OldLocation = OldLocation & "/"
        AddComputerNameToHistory
    Case 1
    ' Example : www.yahoo.com/r
        ' Call the OneSlashURL sub
        OneSlashURL
    Case 2
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/r/m1
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com
            ' Call the TwoSlashURL sub
            TwoSlashURL
        End If
    Case Else
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/homer/?http://greetings.yahoo.com
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com/r/m1
            ' Call the ThreeSlashURL sub
            ThreeSlashURL
        End If
End Select
End Sub




Public Sub AddComputerNameToHistory()
' Error number 35602 is generated when the key is not
' unique. Since the NewLocation (Computer Name) is used
' as a key, the ErrHandler will work like the following:
' if the error number is not 35602, add the NewLocation
' to the HistoryTree. This is easier than assigning a
' different key to each node
On Error GoTo ErrHandler
' If you remove the WebBrowser1.GoBack from the Form_Load
' the NewLocation will be a null string and the
' OldLocation will be"http:///", that's why you will have
' to add "And OldLocation <> "http:///" in the following
' If statement

ErrHandler:
If Err.Number <> 35602 And OldLocation <> "/" Then
Set nodUrl = TreeHistory.Nodes.Add(Today, tvwChild, NewLocation, NewLocation, 2, 3)
' Sort the nodes
nodUrl.Sorted = True
End If
' Call the AddUrlToHistory sub
AddUrlToHistory
End Sub


Public Sub AddUrlToHistory()
On Error GoTo ErrHandler2
ErrHandler2:
If Err.Number <> 35602 And OldLocation <> "/" Then
TreeHistory.Nodes.Add NewLocation, tvwChild, OldLocation, OldLocation, 4
End If
End Sub


Public Sub AddToday()
If dayAdded = False Then
    Set nodCN = TreeHistory.Nodes.Add(, , Today, ThisDayName, 1)
    nodCN.Sorted = True
    nodCN.Expanded = True
    ' Change the value of dayAdded to True to prevent
    ' from adding the today's name to the TreeHistory
    ' again
    dayAdded = True
End If
End Sub

Public Sub DeleteHistory()
' In the LoadHistory sub the KeyNumber is increased by 1
' each time a name of a day is found. If today's name is
' found, the value in KeyNumber will be assigned to
' DayNumber, and the value 1 is assigned to
' TodayInHistory. If there are more days (after today) in
' the history file, the KeyNumber will increase and
' becomes greater than DayNumber.
' Here is an example of how the DeleteHistory sub works:
' if today is Wednesday, and Thursday was found in the
' history file, that means that this is the last week's
' history and it has to be cleared.
' But if today's name was not found in the history file,
' the value of KeyNumber will not be assigned to
' DayNumber (in the LoadHistory sub) which means that the
' value of KeyNumber will be greater than DayNumber and
' the history file will be cleared. To prevent that from
' happening, TodayInHistory is also used in the
' DeleteHistory sub like the following:

If (DayNumber < KeyNumber) And (TodayInHistory = 1) Then
    Open App.Path & "\history.txt" For Output As #4
    Close #4
    ' The TreeHistory must be cleared or else the old
    ' history will still be visible in it
    TreeHistory.Nodes.Clear
    ' dayAdded will be used in the AddToday sub
    dayAdded = False
End If
End Sub



Private Sub TreeHistory_NodeClick(ByVal Node As MSComctlLib.Node)
' This code is used to expand, close and navigate the
' TreeHistory with a single click only

' If Node.Children = 0 that means that the clicked node
' is a URL, so the WebBrowser should navigate it
If Node.Children = 0 Then
    WebBrowser1.Navigate Node.text
Else
' If Node.Children <> 0 that means that the clicked node
' is either a day or a computer name and the WebBrowser
' should not navigate it. If the node is expanded, it
' will be closed; if not, it will be expanded
    If Node.Expanded = True Then
        Node.Expanded = False
    Else
        Node.Expanded = True
    End If
End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
' OldLocation will be used in the urlTest sub
OldLocation = url
' Show the navigated URL in the Address Combo Box
cboAddress.text = url
' The history must store every URL navigated, and not
' just the ones entered in the Address Combo Box, that's
' why the urlTest sub must be called
urlTest
End Sub

'********************************************
'********************************************
'*****     HISTORY  ENDS   HERE      ********
'********************************************
'********************************************
