VERSION 5.00
Begin VB.Form FrmBookmarks 
   BackColor       =   &H00404040&
   Caption         =   "Bookmarks"
   ClientHeight    =   5265
   ClientLeft      =   4560
   ClientTop       =   465
   ClientWidth     =   4950
   Icon            =   "FrmBookmarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   4950
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Or add URL currently visiting"
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
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   4695
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2760
         Picture         =   "FrmBookmarks.frx":1272
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3720
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdHints 
      Caption         =   "&Hints"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame frBookmarks 
      BackColor       =   &H00404040&
      Caption         =   "List of Bookmarks"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      Begin VB.ListBox lstBookmarks 
         Height          =   2010
         ItemData        =   "FrmBookmarks.frx":157C
         Left            =   120
         List            =   "FrmBookmarks.frx":157E
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Add to bookmarks and the press ENTER"
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
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   4695
      Begin VB.TextBox txtBookmarks 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Personal Bookmarks"
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
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmBookmarks.frx":1580
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FrmBookmarks"
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

Private Sub cmdAdd_Click()
Dim strBookmark As String
    'If KeyAscii = 13 Then                                   ' If the user hits enter, it loads into the browser window
        Open App.Path & "\Bookmarks.txt" For Append As #1
        Print #1, FrmInternetBrowser.cboAddress.text
            
            lstBookmarks.AddItem FrmInternetBrowser.cboAddress.text
            'txtBookmarks.text = strBookmark
        Close #1
            FrmInternetBrowser.Toolbar1.Buttons(15).Value = tbrUnpressed
            FrmBookmarks.Visible = False                    ' Unpresses the button in the tool bar and hides the form
            MsgBox "The URL has been saved to your bookmark list.", 8, ""
    'End If
End Sub

Private Sub cmdCancel_Click()
    FrmInternetBrowser.Toolbar1.Buttons(15).Value = tbrUnpressed
    FrmBookmarks.Visible = False                            ' Unpresses the button in the tool bar and hides the form
End Sub

Private Sub cmdHints_Click()
    MsgBox "A few hints on using the bookmarks:" & vbCrLf & vbCrLf & "1)  Accepts either the full URL -> Http:\\www.domain.com" & vbCrLf & "    or shorter URL -> www.domain.com" & vbCrLf & vbCrLf & "2)  The Bookmarks stay on top of IB v1.1 until the button in the toolbar is unpressed", 8, ""
End Sub

Private Sub Form_Load()
    KeepOnTop FrmBookmarks
    
    On Error GoTo FileError
    Open App.Path & "\Bookmarks.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, a$
        lstBookmarks.AddItem a$
    Loop
    Close #1
FileError:
    Open App.Path & "\Bookmarks.list" For Output As #1
    Close #1
End Sub

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmInternetBrowser.Toolbar1.Buttons(15).Value = tbrUnpressed
    FrmBookmarks.Visible = False                            ' Unpresses the button in the tool bar and hides the form
End Sub

Private Sub lstBookmarks_DblClick()
    If InStr(1, "lstBookmarks.text", "Http:\\") = 0 Then
        FrmInternetBrowser.WebBrowser1.Navigate lstBookmarks.text
    Else
        FrmInternetBrowser.WebBrowser1.Navigate "Http:\\" & lstBookmarks.text
    End If
        FrmInternetBrowser.Toolbar1.Buttons(15).Value = tbrUnpressed
        FrmBookmarks.Visible = False                        ' Unpresses the button in the tool bar and hides the form
End Sub

Private Sub lstBookmarks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lstBookmarks.SelCount = 0 Then
        If Button = vbRightButton Then
            Select Case MsgBox("Link" & lstBookmarks.text & " ?", vbYesNo + vbSystemModal, "Fast Links 1.5")
                Case vbYes
                    lstBookmarks.RemoveItem Int(lstBookmarks.ListIndex)
                Case vbNo
            End Select
        End If
    End If
End Sub

Private Sub txtBookmarks_KeyPress(KeyAscii As Integer)
Dim strBookamrk As String
    If KeyAscii = 13 Then                                   ' If the user hits enter, it loads into the browser window
        Open App.Path & "\Bookmarks.txt" For Append As #1
        Print #1, txtBookmarks.text
            lstBookmarks.AddItem txtBookmarks.text
            'txtBookmarks.text = ""
             txtBookmarks.text = strBookamrk
        Close #1
            FrmInternetBrowser.Toolbar1.Buttons(15).Value = tbrUnpressed
            FrmBookmarks.Visible = False                    ' Unpresses the button in the tool bar and hides the form
            MsgBox "The URL has been saved to your bookmark list.", 8, ""
    End If
End Sub
