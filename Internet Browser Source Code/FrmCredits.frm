VERSION 5.00
Begin VB.Form FrmCredits 
   BackColor       =   &H00404040&
   Caption         =   "IB v1.10 - Credits"
   ClientHeight    =   4560
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   7110
   Icon            =   "FrmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7110
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Image Image2 
         Height          =   480
         Left            =   6240
         Picture         =   "FrmCredits.frx":1272
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Written in VB 6.0"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Internet Browser 1.10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "karam_hani@yahoo.com"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   $"FrmCredits.frx":24E4
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
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Programmed by :     Paul Cormie"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   $"FrmCredits.frx":261E
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
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   6615
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6855
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "FrmCredits.frx":26C7
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "paul_cormie@hotmail.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "www.paul_cormie.homestead.com"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdGoThere 
      Caption         =   "&Go There"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCredits"
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
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously.
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
Private Sub cmdOK_Click()
    FrmCredits.Hide                                             ' Hides Credit Form
End Sub

'''''Private Sub cmdGoThere_Click()
'''''        FrmInternetBrowser.cboAddress.text = "http://www.paul_cormie.homestead.com"
'''''        FrmInternetBrowser.WebBrowser1.Navigate FrmInternetBrowser.cboAddress.text
'''''        FrmCredits.Hide
'''''End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCredits.Hide                                             ' Hides Credit Form
End Sub

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                            ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

