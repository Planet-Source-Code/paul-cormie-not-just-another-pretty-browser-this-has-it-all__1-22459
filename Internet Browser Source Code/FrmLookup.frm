VERSION 5.00
Begin VB.Form FrmLookup 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Host Lookup"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "FrmLookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4695
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
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Resolve 
      Caption         =   "&Resolve Host"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
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
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Hostname or IP Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
      Begin VB.TextBox Address 
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
         Height          =   345
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Host Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmLookup"
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
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim Addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0

Private Sub Close_Click()
    Unload Me
End Sub

Private Sub Resolve_Click()
    On Error Resume Next
    Address.text = ""
    If Len(Host.text) = 0 Then
        vbGetHostName
    End If
    vbGetHostByName
End Sub

Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong3, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        Address.text = Trim$(CStr(Asc(IPLong3.Byte4)) + "." + CStr(Asc(IPLong3.Byte3)) _
        + "." + CStr(Asc(IPLong3.Byte2)) + "." + CStr(Asc(IPLong3.Byte1)))
    End If
End Sub

Public Sub vbGetHostName()
    Host = String(64, &H0)
   
    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.text = Host
    End If
End Sub

