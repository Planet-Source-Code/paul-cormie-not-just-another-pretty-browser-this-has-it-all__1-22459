VERSION 5.00
Begin VB.Form FrmTime 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IB v1.10"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FrmTime.frx":0000
   MousePointer    =   4  'Icon
   Picture         =   "FrmTime.frx":030A
   ScaleHeight     =   990
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   840
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, MMMM dd, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   0
      MousePointer    =   11  'Hourglass
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   0
      MouseIcon       =   "FrmTime.frx":0614
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "FrmTime"
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
Private Sub Form_Load()
lblTime.Caption = Time
lblDate.Caption = Date
KeepOnTop FrmTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmPad.Show                                                 ' Toolbar controls
    frmPad.tbToolBar.Buttons(27).Value = tbrUnpressed
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Time
End Sub

Sub KeepOnTop(F As Form)
Const SWP_NOMOVE = 2                                            ' Sets the given form On TopMost
Const SWP_NOSIZE = 1                                            ' See the module for details on how its done
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
