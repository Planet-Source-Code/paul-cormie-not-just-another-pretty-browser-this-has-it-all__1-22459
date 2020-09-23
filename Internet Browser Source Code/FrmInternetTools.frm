VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmInternetTools 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IB v1.10 - Internet Tools"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmInternetTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Exit 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton WhoisBtn 
         Caption         =   "&Whois"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox IPTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox HostTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Ping 
         Caption         =   "&Ping"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Scanner 
         Caption         =   "Port &Scanner"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Source 
         Caption         =   "&View HTML Source Code"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton ListenBtn 
         Caption         =   "Port &Listener"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Lookup 
         Caption         =   "&Host Lookup"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Trace 
         Caption         =   "&Trace Route"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label IPLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Local IP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Host:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmInternetTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ConnectBtn_Click()
Connect.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub FingerBtn_Click()
FingerFrm.Show
End Sub

Private Sub Form_DblClick()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
Dim NameStr As String, SerStr As String
IPTxt.text = Winsock1.LocalIP
HostTxt.text = Winsock1.LocalHostName

Lookup.Enabled = True
ListenBtn.Enabled = True
Scanner.Enabled = True
Trace.Enabled = True
KeepOnTop FrmInternetTools

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub ListenBtn_Click()
FrmListen.Show
End Sub

Private Sub Lookup_Click()
FrmLookup.Show
End Sub

Private Sub Mail_Click()
EmailFrm.Show
End Sub

Private Sub Ping_Click()
FrmPing.Show
End Sub

Private Sub RegIt_Click()
Register.Show
End Sub

Private Sub Scanner_Click()
FrmScanner.lblCurrent = 0
FrmScanner.Show
End Sub

Private Sub Source_Click()
FrmGetHTML.Show
End Sub

Private Sub Speed_Click()
SpeedChk.Show
End Sub

Private Sub Trace_Click()
FrmTrace.Show
End Sub

Private Sub WhoisBtn_Click()
FrmWhois.Show
End Sub

Private Sub Winsck_Click()
Aboutfrm.Show 1
End Sub

