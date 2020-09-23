VERSION 5.00
Begin VB.Form FrmHint 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IB v1.10 - Helpful Hints"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   8310
   Icon            =   "FrmHint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   8310
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      BackColor       =   &H00404040&
      Caption         =   "&Show at Startup"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      BackColor       =   &H00808080&
      Caption         =   "&Next Hint"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   120
      Picture         =   "FrmHint.frx":1272
      ScaleHeight     =   1095
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Helpful Hints..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   7335
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00808080&
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
      Left            =   7200
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "FrmHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ya, Thank the Wizard For this form!


Option Explicit

Dim Tips As New Collection                                                  ' The in-memory database of tips.
Const TIP_FILE = "TIPOFDAY.TXT"                                      ' Name of tips file
Dim CurrentTip As Long                                                       ' Index in collection of tip currently being displayed.

Private Sub DoNextTip()
    ' Select a tip at random.
    'CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If

    ' Show it.
    FrmHint.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  ' Randomize
'    ' Read in the tips file and display a tip at random.
   If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

