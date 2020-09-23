VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form FrmListen 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Listener"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "FrmListen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Stop"
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Listen"
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
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ComboBox cmbPort 
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton optTCP 
      BackColor       =   &H00404040&
      Caption         =   "TCP/IP"
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
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton optUDP 
      BackColor       =   &H00404040&
      Caption         =   "UDP"
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
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Protocol:"
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
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Port:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Status"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
      Begin VB.TextBox txtStatus 
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
      End
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   405
   End
End
Attribute VB_Name = "FrmListen"
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

Private Sub cmdConnect_Click()
    cmdConnect.Enabled = False
    cmbPort.Enabled = False
    cmdDisconnect.Enabled = True
    txtStatus = ""
    If optTCP = True Then
        ws1.Protocol = sckTCPProtocol
    End If
    If optUDP = True Then
        ws1.Protocol = sckUDPProtocol
    End If
    On Error GoTo PortIsOpen
        ws1.Close
        ws1.LocalPort = cmbPort.text
        ws1.Listen
    Exit Sub
PortIsOpen:
    ws1.Close
    If Err.Number = 10048 Then
        txtStatus = "The port " & cmbPort.text & " is already open."
    Else
        txtStatus = "Error: " & Err.Number & vbCrLf & "   " & Err.Description
    End If
    cmdDisconnect.Enabled = False
    cmbPort.Enabled = True
    cmdConnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    ws1.Close
    cmdDisconnect.Enabled = False
    cmbPort.Enabled = True
    cmdConnect.Enabled = True
End Sub

Private Sub Form_Load()
    optTCP = True
End Sub


Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
    If ws1.State <> sckClosed Then ws1.Close
    ws1.Accept (requestID)
    txtStatus.text = "Connection..."
End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    ws1.GetData strData
    txtStatus.text = txtStatus.text & vbCrLf & " - " & strData
End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txtStatus = "Winsock Error: " & Number & vbCrLf & "   " & descriptoin
End Sub
