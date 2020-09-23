VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmTextBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Box"
   ClientHeight    =   3975
   ClientLeft      =   795
   ClientTop       =   585
   ClientWidth     =   12870
   Icon            =   "FrmTextBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   12870
   Begin VB.TextBox txtWordCount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   252
      Left            =   9720
      TabIndex        =   27
      Top             =   600
      Width           =   372
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H000000C0&
      Height          =   252
      Left            =   9720
      TabIndex        =   26
      Top             =   360
      Width           =   372
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5880
      TabIndex        =   25
      Top             =   7560
      Width           =   3492
      Begin VB.CommandButton cmdCancel2 
         BackColor       =   &H8000000C&
         Caption         =   "&Cancel"
         Height          =   372
         Left            =   2280
         MaskColor       =   &H00808080&
         TabIndex        =   17
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdViewSource2 
         Caption         =   "&View Source"
         Height          =   372
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "C&lear Box"
         Height          =   372
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.TextBox Text111111111111111 
      Height          =   2775
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   9480
      Width           =   5412
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6720
      Top             =   5280
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|HTML Files (*.htm, *.html)|*.htm*"
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5880
      TabIndex        =   18
      Top             =   2520
      Width           =   3492
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H8000000C&
         Caption         =   "&Cancel"
         Height          =   372
         Left            =   2280
         MaskColor       =   &H00808080&
         TabIndex        =   14
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdViewSource 
         Caption         =   "&View Source"
         Height          =   372
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear Box"
         Height          =   372
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Frame frTextBox 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5652
      Begin RichTextLib.RichTextBox Text1 
         Height          =   3495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6165
         _Version        =   393217
         TextRTF         =   $"FrmTextBox.frx":1272
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   5880
      TabIndex        =   19
      Top             =   0
      Width           =   3492
      Begin VB.CommandButton cmdLast 
         Caption         =   "Find &Last"
         Height          =   372
         Left            =   2280
         TabIndex        =   5
         Top             =   600
         Width           =   1092
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Find &Next"
         Height          =   372
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   1092
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "Find &First"
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2412
      End
      Begin VB.Label Label1 
         Caption         =   "Search for"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   7200
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save As"
      Filter          =   "Text Files (*.txt)|*.txt|HTML Files (*.htm, *.html)|*.htm*"
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   5880
      TabIndex        =   21
      Top             =   1200
      Width           =   3492
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdLoadFile 
         Caption         =   "L&oad File"
         Height          =   372
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save &As"
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton cmdCut 
         Caption         =   "Cut"
         Height          =   372
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   372
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1092
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblWordCount 
      BackColor       =   &H80000005&
      Caption         =   "Words in Text Box"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10200
      TabIndex        =   29
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblTextCount 
      BackColor       =   &H80000005&
      Caption         =   "Lines in Text Box"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10200
      TabIndex        =   28
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSmall 
      Caption         =   "Click the picture to make Text Box SMALLER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   23
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Image imgBig 
      Height          =   480
      Left            =   5880
      Picture         =   "FrmTextBox.frx":135F
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image imgSmall 
      Height          =   480
      Left            =   5880
      Picture         =   "FrmTextBox.frx":1669
      ToolTipText     =   "Click here to make the text box larger"
      Top             =   8400
      Width           =   480
   End
   Begin VB.Label lblBig 
      Caption         =   "Click the picture to make Text Box BIGGER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   22
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "FrmTextBox"
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
Dim strCount As String
Dim strWordCount As String
Dim str As String
Dim i As Integer
Dim n As Integer
Dim strFileName As String
Dim strFileLoad As String
Dim strMsgBox As String

Private Sub cmdLoadFile_Click()
   'On Error Resume Next
    dlgOpen.ShowOpen                                            ' Opens a common dialog box (open)
        If dlgOpen.FileName = "" Then                           ' If there is nothing in the open file box...how can you open a file?
        Else
            Call LoadText(Text1, dlgOpen.FileName)              ' Calls Loadtext into the textbox
        End If
End Sub

Private Sub cmdRefresh_Click()
    Text1.Text = All
End Sub

Sub LoadText(Lst As textBox, file As String)                    ' GetTextFromFile
    On Error GoTo error                                         ' Call LoadText (Text1,"C:\....Saved.txt")
    Dim mystr As String
    Open file For Input As #1

    Do While Not EOF(1)
        Line Input #1, a$
        texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
    Loop
    Lst = texto$
    Close #1
    Exit Sub
error:                                                          ' Error handleing
    x = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

Private Sub cmdFirst_Click()
    Dim textfound As Integer                                    ' Find first occurence
    cmdNext.Enabled = True                                      ' Enables the cmdFindNext
    Text1.Find (Text2.Text)                                     ' Finds the text in the search box and highlights it,
    Text1.SetFocus                                              ' then sets the focus on the Text1 so the selected text is editable
    textfound = Text1.Find(Text2.Text)                          ' The Text1.find method returns a value of -1 if the searched for text is not found.
    If textfound = -1 Then                                      ' If this is true then it displays a message box.
        MsgBox "End of Document" & vbCr & "Text Not Found", 8, ""
    End If
End Sub

Private Sub cmdNext_Click()                                     ' Find next occurence
    Text1.SetFocus                                              ' Set the focus so the selected text can be directly edited.
    Text1.Find (Text2.Text), Text1.SelStart + 1
End Sub

Private Sub cmdSelectAll_Click()
    Text1.SelStart = 0                                          ' Selects all the text in the current document.
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub cmdLast_Click()
    Text1.SetFocus                                              ' Set the focus so the selected text can be directly edited.
    Text1.Find (Text2.Text), Text1.SelStart - 1                 ' Finds the next instance of the word, starting from the selected text
    cmdFirst.Enabled = False
    cmdNext.Enabled = False
      
End Sub

Private Sub cmdSave_Click()                                     ' Call SaveText (txtQuick,"C:\________temp_file.txt")
    On Error GoTo error
    Dim mystr As String
    Open file For Output As #1
    Print #1, Lst
    Close 1
    MsgBox "The File:" & vbCrLf & vbCrLf & dlgSave.FileName & vbCrLf & vbCrLf & "has been saved", 8, "Confirmation"
    Exit Sub
' Error handleing
error:
    x = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub

Sub SaveText(Lst As RichTextBox, file As String)
    On Error GoTo error                                         ' Call SaveText (txtQuick,"C:\________temp_file.txt")
    Dim mystr As String
    Open file For Output As #1
    Print #1, Lst
    Close 1
    MsgBox "The File:" & vbCrLf & vbCrLf & dlgSave.FileName & vbCrLf & vbCrLf & "has been saved", 8, "Confirmation"
    Exit Sub
error:
    x = MsgBox("There has been a error!", vbOKOnly, "Error")
End Sub

Private Function ReplaceString(MString, SString, RString As String) As String
    Dim a As Integer
    Dim MidString As String
    Dim LeftString As String
    Dim RightString As String
    a = InStr(1, MString, SString)

    If a = 0 Then                                               ' If string not found
        ReplaceString = MString
        Exit Function
    End If
    
    LeftString = Left(MString, a - 1)
    MidString = Left(MString, a + Len(SString))
    a = Len(MString) - Len(MidString)
    RightString = Right(MString, a + 1)
    ReplaceString = LeftString & RString & RightString
    MString = ReplaceString
End Function

Private Sub cmdClear_Click()                                    ' There are two commands due to the resize of the form, 1 visible other not
If Text1.Text = "" Then                                         ' Text1 already clear
    MsgBox " Already Clear ", 8, ""
Else
    strMsgBox = MsgBox(" Are you sure you want to clear the text box?", vbYesNo, "")
    If strMsgBox = vbNo Then
    Else
        For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is textBox Then
                Me.Controls(i).Text = ""
                txtWordCount = ""                               ' Clears the box
                lblTextCount.Visible = False                    ' Hides the text count
                lblWordCount.Visible = False                    ' Hides the word count
            End If
        Next i
    End If
End If
End Sub

Private Sub cmdClear2_Click()                                   ' There are two commands due to the resize of the form, 1 visible other not
If Text1.Text = "" Then                                         ' Text1 already clear
    MsgBox " Already Clear ", 8, ""
Else
    strMsgBox = MsgBox(" Are you sure you want to clear the text box?", vbYesNo, "")
    If strMsgBox = vbNo Then
    Else
        For i = 1 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is textBox Then
                Me.Controls(i).Text = ""
                txtWordCount = ""                               ' Clears the box
                lblTextCount.Visible = False                    ' Hides the text count
                lblWordCount.Visible = False                    ' Hides the word count
            End If
        Next i
    End If
End If
End Sub

'Private Sub cmdViewSource_Click()                               ' There are two commands due to the resize of the form, 1 visible other not
'    Call GetTextFromFile("C:\boot.ini", Text1)
'    lblTextCount.Visible = True
'    lblWordCount.Visible = True
'        strCount = countLines(Text1)
'    txtCount = strCount
'        strWordCount = countWords(Text1)
'    txtWordCount = strWordCount
'        strFileName = dlgOpen.FileName
'    txtFileName = strFileName
'End Sub
Private Sub cmdViewSource_Click()
   On Error Resume Next
    dlgOpen.ShowOpen
        If dlgOpen.FileName = "" Then
        Else
            Call LoadText(Text1, dlgOpen.FileName)
        End If
End Sub





Private Sub cmdViewSource2_Click()                               ' There are two commands due to the resize of the form, 1 visible other not
    ' not yet defined
End Sub

Private Sub cmdCancel_Click()                                   ' There are two commands due to the resize of the form, 1 visible other not
    FrmTextBox.Hide                                             ' Hides This form
End Sub

Private Sub cmdCancel2_Click()                                  ' There are two commands due to the resize of the form, 1 visible other not
    FrmTextBox.Hide                                             ' Hides This form
End Sub

Private Sub Form_Load()
    KeepOnTop FrmTextBox                                        ' See Sub below, but basically keeps the form on top of the Browser
    FrmTextBox.Height = 4335                                    ' Sizes the form
    
    frTextBox.Height = 3855                                     ' These all are to with changing the size
    frTextBox.Top = 0                                           ' of the textbox, frame, and the pictures,
    frTextBox.Left = 120                                        ' buttons, and lables visiblilty
    
    Text1.Height = 3500
    Text1.Top = 240
    Text1.Left = 120
    
    imgSmall.Visible = False
    lblSmall.Visible = False
    
    lblTextCount.Visible = False
    lblWordCount.Visible = False
    
    imgBig.Visible = True
    lblBig.Visible = True
  
    Frame2.Visible = True
    Frame4.Visible = False
    
    If Text1.Text = "" Then                                     ' Diables the find / next / last buttons
        cmdNext.Enabled = False                                 ' If there is nothing in the box
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
    Else
        cmdNext.Enabled = True
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
    End If
End Sub

Sub KeepOnTop(F As Form)
    Const SWP_NOMOVE = 2                                        ' Sets the given form On TopMost
    Const SWP_NOSIZE = 1

    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear                                             ' Delete everthing in the clipboard
    Clipboard.SetText Text1.SelText, 1                          ' Put your text into it on place 1
End Sub

Private Sub cmdCut_Click()
    Clipboard.Clear                                             ' Delete everthing in the clipboard
    Clipboard.SetText Text1.SelText, 1                          ' Put your text into it on place 1
    Text1.SelText = ""                                          ' Delete everyting that was selected in the textbox
End Sub

Private Sub cmdPaste_Click()
    Text1.SelText = Clipboard.GetText(1)                        ' get the text in the clipboard on place 1 and
                                                                ' Place it on the selected area in the textbox
                                                                ' If nothing is selected it will be place on the place of writing cursor
End Sub

Private Sub Form_Unload(cancel As Integer)
    FrmTextBox.Hide
    FrmInternetBrowser.Show
End Sub

Private Sub imgBig_Click()                                      ' See form Load
    FrmTextBox.Height = 9400
    
    frTextBox.Height = 8895
    frTextBox.Top = 0
    frTextBox.Left = 120
    
    Text1.Height = 8535
    Text1.Top = 240
    Text1.Left = 120
    
    imgBig.Visible = False
    lblBig.Visible = False
    
    imgSmall.Visible = True
    lblSmall.Visible = True
    
    Frame2.Visible = False
    Frame4.Visible = True
    
End Sub

Private Sub imgSmall_Click()                                    ' See form Load
    FrmTextBox.Height = 4335
    
    frTextBox.Height = 3855
    frTextBox.Top = 0
    frTextBox.Left = 120
    
    Text1.Height = 3500
    Text1.Top = 240
    Text1.Left = 120
    
    imgSmall.Visible = False
    lblSmall.Visible = False
    
    imgBig.Visible = True
    lblBig.Visible = True
    
    Frame2.Visible = True
    Frame4.Visible = False
    
End Sub

Private Sub Text1_Change()
        cmdNext.Enabled = True                                  ' Enables the buttons once the text has been changed
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
End Sub
