VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmGetHTML 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get HTML"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "FrmGetHTML.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1800
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtURL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "URL:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6735
      Begin RichTextLib.RichTextBox HTML 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4048
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"FrmGetHTML.frx":1272
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   5760
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton GetIt 
      Caption         =   "&View Source"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "FrmGetHTML"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

'Get HTML
Private Sub GetIt_Click()
    Dim Strsource As String
    Strsource = Inet1.OpenURL(txtURL.text)
    HTML.text = Strsource
End Sub


