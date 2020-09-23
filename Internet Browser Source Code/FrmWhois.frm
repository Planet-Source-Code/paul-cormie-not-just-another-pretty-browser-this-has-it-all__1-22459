VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmWhois 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whois"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "FrmWhois.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Close 
      Caption         =   "&Cancel"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdWhois 
      Caption         =   "&Whois"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ComboBox ServerTxt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmWhois.frx":1272
      Left            =   1080
      List            =   "FrmWhois.frx":1274
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5175
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   4080
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtWhois 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Server:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Domain:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmWhois"
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
Private Sub Close_Click()
    Unload Me
End Sub

Private Sub cmdWhois_Click()
    Winsock1.Close
    Dim WhoisStr As String
    txtWhois.text = ""
    Winsock1.Connect ServerTxt, 43
End Sub

Private Sub Form_Load()
    ServerTxt.AddItem "198.41.0.8"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next
    Winsock1.SendData ("whois " & Host.text & vbCrLf)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dataA
    Winsock1.GetData dataA, vbString
    txtWhois.text = txtWhois.text & dataA '& vbCrLf
    Dim counter As Long
    counter = 1
start:
       Dim Search, where   ' Declare variables.
       ' Get search string from user.
       Search = Chr$(10)
       where = InStr(counter, txtWhois.text, Search, vbTextCompare) ' Find string in text.
       'MsgBox Where
       If where Then   ' If found,
          txtWhois.SelStart = where - 1   ' set selection start and
          txtWhois.SelLength = Len(Search)
          txtWhois.SelText = vbCrLf
          counter = where + txtWhois.SelLength + 2 ': 'MsgBox counter
       Else
          Exit Sub  ' Notify user.
       End If

GoTo start
    Winsock1.Close
End Sub
