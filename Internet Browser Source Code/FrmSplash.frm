VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmSplash 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5895
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000011&
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00808080&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00808080&
      Caption         =   "&OK"
      Height          =   375
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H00808080&
      Caption         =   "C&redits"
      Height          =   375
      Left            =   5760
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5640
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3255
      Begin VB.TextBox txtInternetPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtComputerName 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Internet Port:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Computer Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   " IP address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   6120
      Top             =   6720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000A&
      Height          =   4770
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
         Height          =   3375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   9495
         _cx             =   22888812
         _cy             =   22878017
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         Stacking        =   "below"
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2520
         Picture         =   "FrmSplash.frx":1272
         Stretch         =   -1  'True
         Top             =   120
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   3615
         Left            =   120
         Top             =   720
         Width           =   9735
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Version 1.10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6840
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Internet Browser"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   3480
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Suggested Resolution : 1024 x 768"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      Top             =   5160
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "This application works best under "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   3600
      TabIndex        =   15
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lblPlatform 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Windows® 95/98/ME/NT/2000"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   3600
      TabIndex        =   14
      Top             =   5520
      Width           =   1890
   End
End
Attribute VB_Name = "frmSplash"
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
Option Explicit

Private Sub cmdCancel_Click()
    Dim strMsgBox As String                                 ' You sure you want to quit?  Really?
    strMsgBox = MsgBox(" Are you sure you want to exit Internet Browser 1.1? ", vbOKCancel, "Exit")
                    If strMsgBox = "2" Then
                    Else
                        End
                    End If
End Sub

Private Sub cmdCredits_Click()
    FrmCredits.Show                                         ' Shows the Credits Form
End Sub

Private Sub cmdOK_Click()
    Unload Me                                               ' If you click OK it loads the Browser
    FrmInternetBrowser.Show
    FrmCredits.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me                                               ' If you press enter it loads the Browser
    FrmInternetBrowser.Show
    FrmCredits.Hide
End Sub

Private Sub Form_Load()
    Flash1.Movie = App.Path & "\intro.swf"                  ' Needed for the FLASH intro
    
    Winsock1.LocalPort = 80                                 ' Needed for the computer information
    Winsock1.Listen
    DoEvents
    Me.Show
    DoEvents
    txtIP.text = Winsock1.LocalIP                           ' Displays your IP address
    txtComputerName.text = Winsock1.LocalHostName           ' Computer Name
    txtInternetPort.text = Winsock1.LocalPort               ' Port number

End Sub

Private Sub Frame1_Click()
    Unload Me                                               ' If you click anywhere on the Frame it loads the Browser
    FrmInternetBrowser.Show
    FrmCredits.Hide
End Sub

Private Sub Timer1_Timer()
    Unload Me                                               ' 10 seconds until it loads the browser automatically
    FrmInternetBrowser.Show
    FrmCredits.Hide
End Sub

Private Sub Flash1_FSCommand(ByVal command As String, ByVal args As String)
    Select Case command                                     ' Needed for the FLASH intro on the splash
    End Select
End Sub

