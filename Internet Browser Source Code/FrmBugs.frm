VERSION 5.00
Begin VB.Form FrmBugs 
   BackColor       =   &H00404040&
   Caption         =   "IB v1.10 - Bugs"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
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
   Icon            =   "FrmBugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   $"FrmBugs.frx":1272
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "about:blank/"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "about:blank"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Monday"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "IB v1.10 Bugs"
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
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "History Bug - If, on application load, an error occurs, clear the history except for one day.  For example enter in history.txt :"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmBugs"
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
Private Sub cmdOK_Click()
    FrmBugs.Hide
End Sub

