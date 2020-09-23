VERSION 5.00
Begin VB.Form FrmPing 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "FrmPing.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmTop 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.HScrollBar ScrollPacket 
         Height          =   255
         Left            =   1440
         Max             =   100
         Min             =   1
         TabIndex        =   3
         Top             =   1080
         Value           =   1
         Width           =   3975
      End
      Begin VB.HScrollBar ScrollTimes 
         Height          =   255
         Left            =   1440
         Max             =   10
         Min             =   1
         TabIndex        =   2
         Top             =   720
         Value           =   1
         Width           =   3975
      End
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
         Left            =   1440
         TabIndex        =   1
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label lblPacketSize 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "32"
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
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblPacket 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Packet:"
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
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblPingTimes 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "1"
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
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   105
      End
      Begin VB.Label lblPings 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Ping(s):"
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
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblIpHost 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Ip/Host:"
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
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   705
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
      Left            =   4680
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "&Ping"
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
      Left            =   3240
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame frmBottom 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   5655
      Begin VB.TextBox txtStatus 
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "FrmPing"
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
Dim PingTimes As Integer
Dim Speed As Long
Dim IP As String
Dim KeepGoing As Integer
Dim TotalNum As Long
Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
Dim ExitTheFor As Integer
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

Private Sub cmdClose_Click()
    FrmPing.Hide
End Sub

Private Sub cmdPing_Click()
    Speed = 0
    PingTimes = 0
    cmdPing.Enabled = False
    ScrollTimes.Enabled = False
    ScrollPacket.Enabled = False
    txtStatus = ""
    szBuffer = Space(Val(lblPacketSize))
    vbWSAStartup
    If Len(Host.text) = 0 Then
        vbGetHostName
    End If
        vbGetHostByName
        vbIcmpCreateFile
        pIPo2.TTL = Trim$(255)
        '
    For Times = 1 To lblPingTimes
    If ExitTheFor = 1 Then ExitTheFor = 0: Exit For
        vbIcmpSendEcho
    Next
        vbIcmpCloseHandle
        vbWSACleanup
        ScrollTimes.Enabled = True
        ScrollPacket.Enabled = True
        cmdPing.Enabled = True
        
    On Error GoTo skipit
        Speed = Speed / PingTimes
        txtStatus = txtStatus & vbCrLf & " Average Speed: " & Speed & "."
        txtStatus.SelStart = Len(txtStatus)
    Exit Sub
skipit:
End Sub

Public Sub GetRCode()
RCode = ""
    If pIPe.Status = 0 Then RCode = "Success"
    If pIPe.Status = 11001 Then RCode = "Buffer too Small"
    If pIPe.Status = 11002 Then RCode = "Destination Unreahable"
    If pIPe.Status = 11003 Then RCode = "Dest Host Not Reachable"
    If pIPe.Status = 11004 Then RCode = "Dest Protocol Not Reachable"
    If pIPe.Status = 11005 Then RCode = "Dest Port Not Reachable"
    If pIPe.Status = 11006 Then RCode = "No Resources Available"
    If pIPe.Status = 11007 Then RCode = "Bad Option"
    If pIPe.Status = 11008 Then RCode = "Hardware Error"
    If pIPe.Status = 11009 Then RCode = "Packet too Big"
    If pIPe.Status = 11010 Then RCode = "Reqested Timed Out"
    If pIPe.Status = 11011 Then RCode = "Bad Request"
    If pIPe.Status = 11012 Then RCode = "Bad Route"
    If pIPe.Status = 11014 Then RCode = "TTL Exprd Reassemb"
    If pIPe.Status = 11015 Then RCode = "Parameter Problem"
    If pIPe.Status = 11016 Then RCode = "Source Quench"
    If pIPe.Status = 11017 Then RCode = "Option too Big"
    If pIPe.Status = 11018 Then RCode = "Bad Destination"
    If pIPe.Status = 11019 Then RCode = "Address Deleted"
    If pIPe.Status = 11020 Then RCode = "Spec MTU Change"
    If pIPe.Status = 11021 Then RCode = "MTU Change"
    If pIPe.Status = 11022 Then RCode = "Unload"
    If pIPe.Status = 11050 Then RCode = "General Failure"

    DoEvents

        If RCode <> "" Then
            If RCode = "Success" Then
                Speed = Speed + Val(Trim$(CStr(pIPe2.RoundTripTime)))
                txtStatus.text = txtStatus.text + " Reply from " + RespondingHost + ": Bytes = " + Trim$(CStr(pIPe2.DataSize)) + " RTT = " + Trim$(CStr(pIPe2.RoundTripTime)) + "ms TTL = " + Trim$(CStr(pIPe2.Options.TTL)) + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            Exit Sub
            End If
            KeepGoing = 1
            txtStatus.text = txtStatus.text & RCode
        Else
            KeepGoing = 1
            txtStatus.text = txtStatus.text & RCode
        End If
        txtStatus.SelStart = Len(txtStatus)
    End Sub

Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong2, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong2.Byte4)) + "." + CStr(Asc(IPLong2.Byte3)) _
        + "." + CStr(Asc(IPLong2.Byte2)) + "." + CStr(Asc(IPLong2.Byte1)))
    End If
End Sub

Public Sub vbGetHostName()
    
    Host = String(64, &H0)
    
    If gethostname(Host, HostLen) = SOCKET_ERROR Then
        sMsg = "WSock32 Error" & Str$(WSAGetLastError())
        txtStatus = sMsg
        ExitTheFor = 1
    Else
        Host = Left$(Trim$(Host), Len(Trim$(Host)) - 1)
        Host.text = Host
    End If
End Sub

Public Sub vbIcmpSendEcho()
    Dim NbrOfPkts As Integer
    For NbrOfPkts = 1 To Trim$(1)

        DoEvents
            bReturn = IcmpSendEcho(hIP, Addr, szBuffer, Len(szBuffer), pIPo2, pIPe2, Len(pIPe2) + 8, 2700)
            If bReturn Then
                If KeepGoing = 1 Then KeepGoing = 0: Exit For
                PingTimes = PingTimes + 1
                RespondingHost = CStr(pIPe2.Address(0)) + "." + CStr(pIPe2.Address(1)) + "." + CStr(pIPe2.Address(2)) + "." + CStr(pIPe2.Address(3))
                GetRCode
            Else
                txtStatus.text = txtStatus.text + " Request Timeout" + vbCrLf
                txtStatus.SelStart = Len(txtStatus)
            End If
        Next NbrOfPkts
    End Sub

Sub vbWSAStartup()
Dim wsdaata As WSAdata
    iReturn = WSAStartup(&H101, WSAdata)


    If iReturn <> 0 Then ' If WSock32 error, then tell me about it
        txtStatus = "WSock32.dll is Not responding!"
        ExitTheFor = 1
    End If


    If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(Str$(HiByte(WSAdata.wVersion)))
        sLowByte = Trim$(Str$(LoByte(WSAdata.wVersion)))
        sMsg = "WinSock Version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is Not supported "
        txtStatus = sMsg
        ExitTheFor = 1
        End
    End If


    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            txtStatus = sMsg
            ExitTheFor = 1
        End
    End If
    
    MaxSockets = WSAdata.iMaxSockets


    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = WSAdata.iMaxUdpDg


    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    
    Description = ""


    For i = 0 To WSADESCRIPTION_LEN
        If WSAdata.szDescription(i) = 0 Then Exit For
        Description = Description + Chr$(WSAdata.szDescription(i))
    Next i
    Status = ""


    For i = 0 To WSASYS_STATUS_LEN
        If WSAdata.szSystemStatus(i) = 0 Then Exit For
        Status = Status + Chr$(WSAdata.szSystemStatus(i))
    Next i
End Sub

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function

Public Sub vbWSACleanup()
    iReturn = WSACleanup()
End Sub

Public Sub vbIcmpCloseHandle()
    bReturn = IcmpCloseHandle(hIP)
End Sub

Public Sub vbIcmpCreateFile()
    hIP = IcmpCreateFile()
End Sub

Private Sub Form_Load()
    ScrollPacket.Value = 32
    vbWSAStartup
    vbWSACleanup
End Sub

Private Sub ScrollPacket_Change()
    lblPacketSize = ScrollPacket.Value
End Sub

Private Sub ScrollTimes_Change()
    lblPingTimes = ScrollTimes.Value
End Sub

