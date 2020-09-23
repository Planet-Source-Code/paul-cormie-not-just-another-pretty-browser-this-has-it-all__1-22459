VERSION 5.00
Begin VB.Form FrmSearch 
   BackColor       =   &H00404040&
   Caption         =   "Search"
   ClientHeight    =   2910
   ClientLeft      =   2460
   ClientTop       =   1545
   ClientWidth     =   4965
   Icon            =   "FrmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4965
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      Begin VB.ListBox lstSearch 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "FrmSearch.frx":1272
         Left            =   120
         List            =   "FrmSearch.frx":129A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FrmSearch.frx":131E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Search Engines"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmSearch"
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
Private Sub cmdHints_Click()
        MsgBox "A few hints on using the Search:" & vbCrLf & vbCrLf & "1)  Double Click on name of search engine to load page" & vbCrLf & vbCrLf & "2)  The Search stays on top of IB v1.1 until the button in the toolbar button is unpressed", 8, ""
End Sub

Private Sub Form_Load()
    KeepOnTop FrmSearch                                         ' Obviously keeps the search form on top of the browser
End Sub

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                            ' Sets the given form On TopMost
Const SWP_NOSIZE = 1                                            ' See the module for details on how its done
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmInternetBrowser.Toolbar1.Buttons(13).Value = tbrUnpressed
    FrmSearch.Visible = False                                    ' Unpresses the button in the tool bar and hides the form
End Sub

Private Sub lstSearch_DblClick()
' These values are in the listbox (lstSearch) including the spaces (for formating)
' basically when the selected is what is in the box, goto...


If lstSearch.text = "AltaVista" Then
    Go = "http://www.altavista.com"
End If

If lstSearch.text = "Yahoo" Then
    Go = "http://www.yahoo.com"
End If

If lstSearch.text = "Google" Then
    Go = "http://www.google.com"
End If

If lstSearch.text = "WebCrawler" Then
    Go = "http://www.webcrawler.com"
End If

If lstSearch.text = "Excite" Then
    Go = "http://search.excite.com"
End If

If lstSearch.text = "Go.com" Then
    Go = "http://www.goto.com"
End If

If lstSearch.text = "Hotbot" Then
Go = "http://www.hotbot.com"
End If

If lstSearch.text = "Lycos" Then
Go = "http://www.Lycos.com"
End If

If lstSearch.text = "MSN" Then
Go = "http://www.msn.com/"
End If

If lstSearch.text = "Direct Hit" Then
Go = "http://www.directhit.com/"
End If

If lstSearch.text = "Netscape Netcenter" Then
Go = "http://www.netcenter.com/"
End If

' This is a SWEEEET search engine...better than the web based ones
If lstSearch.text = "Or Download Copernic" Then
Go = "http://www.copernic.com/download/"
End If

FrmInternetBrowser.WebBrowser1.Navigate Go                      ' loads the URL into the browser
FrmInternetBrowser.Toolbar1.Buttons(13).Value = tbrUnpressed    ' unpresses the search button
FrmSearch.Visible = False

End Sub
