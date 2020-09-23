VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IB v 1.10"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "FrmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox lstFiles 
      Height          =   2115
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3120
      Top             =   3000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      Begin ComctlLib.Slider slVolume 
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   2
         Max             =   300
         SelStart        =   150
         TickStyle       =   3
         Value           =   150
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
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
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
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
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtSong 
         BackColor       =   &H8000000B&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
      Begin VB.Line Line3 
         X1              =   5145
         X2              =   4905
         Y1              =   330
         Y2              =   210
      End
      Begin VB.Line Line2 
         X1              =   4920
         X2              =   5160
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         X1              =   4905
         X2              =   4905
         Y1              =   210
         Y2              =   330
      End
      Begin VB.Label lblHide 
         Caption         =   ">>"
         Height          =   210
         Left            =   6240
         TabIndex        =   15
         ToolTipText     =   "Show/Hide files"
         Top             =   720
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox lstSelected 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.ListBox lstFiles2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ComboBox cmbDrives 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '*********************************************************
    '* This program was created by andreas gustafsson.       *
    '* Please do not change/remove this text                 *                               *
    '* Feel free to edit the code as you wish                *
    '* send comments to andreasgustafsson1@hotmail.com       *
    '* References: Windows media player,                     *
    '*             Microsoft scripting runtime               *
    '*Components: Windows common controls 5.0                *
    '*********************************************************
    Option Explicit
    'Dim Fso As New FileSystemObject
    'Dim Fso As New F
    Dim Player As New MediaPlayer.MediaPlayer
    'The selected drive
    Dim strDrive As String
    'The folderpath
    Dim strFolder As String
    'Collection to store the selected filepaths
    Dim Col As New Collection
    Dim Playing As Integer
    Dim Volume As Integer

Private Sub cmbDrives_Click()
    Dim drive As drive
    Dim File As File
    Dim SubFolder As Folder
    Dim i As Integer
    i = 0
    lstFiles.Clear
    If cmbDrives = "" Then Exit Sub
    strDrive = cmbDrives.text
    strFolder = ""
    Set drive = Fso.GetDrive(cmbDrives.text)
    If drive.IsReady Then
        For Each File In drive.RootFolder.Files
            If InStr(File, "mp3") Then
                lstFiles.AddItem File.Name, i
                i = i + 1
            End If
            
        Next
        i = lstFiles.ListCount
        For Each SubFolder In drive.RootFolder.SubFolders
        lstFiles.AddItem SubFolder, i
        i = i + 1
        Next
    Else
        MsgBox "Drives not ready"
    End If
End Sub
'adds the selected files to the list amd collection
Private Sub cmdAdd_Click()
    Dim i As Integer
    Dim J As Integer
    If Col.Count > 0 Then i = Col.Count
    If InStr(lstFiles.text, ":\") Then Exit Sub
    For J = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(J) Then
            Col.Add strDrive & strFolder & "\" & lstFiles.List(J), CStr(i)
            lstSelected.AddItem lstFiles.List(J), i
            i = i + 1
        End If
    Next J
End Sub
'Play the previous song
Private Sub cmdBack_Click()
    If Col.Count = 0 Then Exit Sub
    Player.Stop
    If Playing <> 1 Then Playing = Playing - 1
    Player.Open Col.Item(Playing)
    Timer1.Enabled = True
    txtSong = Right(Col.Item(Playing), Len(Col.Item(Playing)) - InStrRev(Col.Item(Playing), "\", , vbTextCompare))
    Me.Caption = "Player - " & txtSong
End Sub
'Clear the collection and list
Private Sub cmdClear_Click()
    Dim i As Integer
    i = Col.Count
    While i > 0
        Col.Remove i
        i = i - 1
    Wend
    lstSelected.Clear
End Sub
'Play next song in the collection
Private Sub cmdNext_Click()
    Player.Stop
    If Col.Count = 0 Then Exit Sub
    If Playing < Col.Count Then Playing = Playing + 1
    Player.Open Col.Item(Playing)
    Timer1.Enabled = True
    txtSong = Right(Col.Item(Playing), Len(Col.Item(Playing)) - InStrRev(Col.Item(Playing), "\", , vbTextCompare))
    Me.Caption = "Player - " & txtSong
End Sub

Private Sub cmdPlay_Click()
    Playing = 1
    If Player.PlayState = mpPaused Then
        Player.Play
    Else
        If Col.Count = 0 Then Exit Sub
        Player.Open Col.Item(Playing)
    End If
    Timer1.Enabled = True
    txtSong = Right(Col.Item(Playing), Len(Col.Item(Playing)) - InStrRev(Col.Item(Playing), "\", , vbTextCompare))
    Me.Caption = "Player - " & txtSong
    Locking False
End Sub
Private Sub Locking(State As Boolean)
    cmdAdd.Enabled = State
    cmdRemove.Enabled = State
    cmdClear.Enabled = State
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
Back:
    If lstSelected.SelCount > 0 Then
        For i = 0 To lstSelected.ListCount - 1
            If lstSelected.Selected(i) Then
                Col.Remove i + 1
                lstSelected.RemoveItem (i)
                GoTo Back
            End If
        Next i
    End If
End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
    Player.Stop
    Playing = 0
    Locking True
    txtSong = ""
    Me.Caption = "Player"
End Sub
'Moves to the parent folder (if any)
Private Sub cmdup_Click()
    Dim Folder As Folder
    Dim File As File
    Dim SubFolder As Folder
    Dim i As Integer
    If strDrive = "" Then Exit Sub
    Set Folder = Fso.GetFolder(strDrive & strFolder)
    strFolder = Left(strFolder, InStr(strFolder, "\"))
        lstFiles.Clear
        If Not Folder.ParentFolder Is Nothing Then
            For Each File In Folder.ParentFolder.Files
                If InStr(File, "mp3") Then
                    lstFiles.AddItem File.Name, i
                    i = i + 1
                End If
            Next
        
            i = lstFiles.ListCount
            For Each SubFolder In Folder.ParentFolder.SubFolders
                lstFiles.AddItem SubFolder, i
                i = i + 1
            Next
        Else
            For Each File In Folder.Files
                If InStr(File, "mp3") Then
                    lstFiles.AddItem File.Name, i
                    i = i + 1
                End If
                
            Next
            i = lstFiles.ListCount
            For Each SubFolder In Folder.SubFolders
                lstFiles.AddItem SubFolder, i
                i = i + 1
            Next
        End If
End Sub

Private Sub cmdPause_Click()
    If Col.Count = 0 Then Exit Sub
    'if its paused
    If Player.PlayState = mpPaused Then
        Player.Play
        Timer1.Enabled = True
    Else
        Player.Pause
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim drive 'As drive
    Dim i As Integer
    i = 0
    Me.Height = 1410
    Me.Width = 6870
    Volume = Player.Volume
    slVolume_Scroll
    For Each drive In Fso.Drives
        cmbDrives.AddItem drive.Path, i
        i = i + 1
    Next
End Sub
'Change the color of the label if the mouse move outside the label
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHide.ForeColor = &H80000012
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Player.Stop
    Set Player = Nothing
    Set Col = Nothing
    Set Fso = Nothing
End Sub

Private Sub lblHide_Click()
    If Me.Height = 3855 Then
        Me.Height = 1410
    Else
        Me.Height = 3855
    End If
    Me.Width = 6870
End Sub

Private Sub lblHide_DblClick()
    If Me.Height = 3855 Then
        Me.Height = 1410
    Else
        Me.Height = 3855
    End If
    Me.Width = 6870
End Sub
'Change the color when the mouse is over the label
Private Sub lblHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHide.ForeColor = &H80000011
End Sub

Private Sub lstFiles_Click()
    Dim Folder As Folder
    Dim SubFolder As Folder
    Dim File As File
    Dim i As Integer
    i = 0
    If Not lstFiles.SelCount > 1 Then
        'if its a folder
        If InStr(lstFiles.text, ":\") Then
            Set Folder = Fso.GetFolder(lstFiles.text)
            lstFiles.Clear
            strFolder = strFolder & "\" & Folder.Name
            'Add all .mp3 files
            For Each File In Folder.Files
                If InStr(File, ".mp3") Then
                    lstFiles.AddItem File.Name, i
                    i = i + 1
                End If
            Next
            i = lstFiles.ListCount
            'Add subfolders
            For Each SubFolder In Folder.SubFolders
                lstFiles.AddItem SubFolder, i
                i = i + 1
            Next
        End If
    End If
End Sub
'Start playing the selected song
Private Sub lstSelected_DblClick()
    Dim i As Integer
    For i = 0 To lstSelected.ListCount - 1
        If lstSelected.Selected(i) Then
            Player.Stop
            Playing = i + 1
            Player.Open Col.Item(Playing)
            Timer1.Enabled = True
            txtSong = Right(Col.Item(Playing), Len(Col.Item(Playing)) - InStrRev(Col.Item(Playing), "\", , vbTextCompare))
            Me.Caption = "Player - " & txtSong
            Locking False
        End If
    Next i
End Sub

Private Sub slVolume_Scroll()
    Player.Volume = Volume * (1 + (300 - slVolume.Value) / 100)
End Sub
'Check if the plater stopped
Private Sub Timer1_Timer()
    If Player.PlayState = mpStopped Then
        Playing = Playing + 1
        If Playing > Col.Count Then
            Timer1.Enabled = False
            Exit Sub
        End If
        Player.Open Col.Item(Playing)
        txtSong = Right(Col.Item(Playing), Len(Col.Item(Playing)) - InStrRev(Col.Item(Playing), "\", , vbTextCompare))
        Me.Caption = "Player - " & txtSong
    End If
End Sub
