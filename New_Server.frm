VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_NewServer 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9525
   ClientLeft      =   255
   ClientTop       =   735
   ClientWidth     =   10500
   ControlBox      =   0   'False
   Icon            =   "New_Server.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "New_Server.frx":030A
   Picture         =   "New_Server.frx":0614
   ScaleHeight     =   9525
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command71_ 
      Caption         =   "&Connect"
      Height          =   300
      Left            =   9645
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4700
      Left            =   15
      Picture         =   "New_Server.frx":1E76F
      ScaleHeight     =   4695
      ScaleWidth      =   10065
      TabIndex        =   18
      Top             =   4710
      Visible         =   0   'False
      Width           =   10065
      Begin VB.ListBox IPList 
         Height          =   2400
         Left            =   7710
         TabIndex        =   32
         Top             =   2235
         Width           =   1650
      End
      Begin VB.ListBox UserList_ 
         Height          =   1425
         Left            =   7710
         TabIndex        =   31
         Top             =   360
         Width           =   1650
      End
      Begin VB.TextBox Text6_ 
         Height          =   300
         Left            =   75
         MaxLength       =   200
         TabIndex        =   28
         Top             =   4335
         Width           =   7515
      End
      Begin VB.TextBox Text5_ 
         BorderStyle     =   0  'None
         Height          =   2550
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1635
         Width           =   7395
      End
      Begin VB.TextBox Text3_ 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   180
         TabIndex        =   22
         Text            =   "127.0.0.1"
         Top             =   330
         Width           =   2535
      End
      Begin VB.TextBox Text4_ 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1950
         TabIndex        =   21
         Text            =   "20000"
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox Text2_ 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   20
         Text            =   "Password"
         Top             =   975
         Width           =   1530
      End
      Begin VB.TextBox Text1_ 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3015
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   330
         Width           =   2130
      End
      Begin VB.Image Command9 
         Height          =   330
         Left            =   4905
         Picture         =   "New_Server.frx":3C8CA
         Top             =   990
         Width           =   1830
      End
      Begin VB.Image Command7_ 
         Height          =   330
         Left            =   5235
         Picture         =   "New_Server.frx":3CE2E
         Top             =   525
         Width           =   2430
      End
      Begin VB.Image Image4 
         Height          =   330
         Left            =   6600
         MouseIcon       =   "New_Server.frx":3D3E5
         Picture         =   "New_Server.frx":3D6EF
         Top             =   1000
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   330
         Left            =   5325
         Picture         =   "New_Server.frx":3DB9B
         Top             =   60
         Width           =   2160
      End
      Begin VB.Label Label6_ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Closed."
         Height          =   255
         Left            =   3750
         TabIndex        =   30
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connection"
         Height          =   600
         Left            =   3675
         TabIndex        =   35
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP List :"
         Height          =   255
         Left            =   7710
         TabIndex        =   34
         Top             =   1935
         Width           =   1650
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User List :"
         Height          =   255
         Left            =   7710
         TabIndex        =   33
         Top             =   60
         Width           =   1650
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Chat Window :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   75
         TabIndex        =   29
         Top             =   1395
         Width           =   7515
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server Hostname / IP Address :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   75
         TabIndex        =   26
         Top             =   60
         Width           =   2730
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Server Port :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1875
         TabIndex        =   25
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   75
         TabIndex        =   24
         Top             =   690
         Width           =   1650
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Username  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2955
         TabIndex        =   23
         Top             =   60
         Width           =   2310
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   47005
      Left            =   10005
      Picture         =   "New_Server.frx":3E206
      ScaleHeight     =   47010
      ScaleWidth      =   10110
      TabIndex        =   36
      Top             =   4710
      Visible         =   0   'False
      Width           =   10110
      Begin VB.CommandButton Command1 
         Caption         =   "Hide Help"
         Height          =   315
         Left            =   3750
         TabIndex        =   38
         Top             =   4335
         Width           =   1800
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   3765
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Text            =   "New_Server.frx":5C361
         Top             =   435
         Width           =   7695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   2925
      Top             =   30
   End
   Begin VB.Timer Timer3 
      Left            =   1965
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   17
      Top             =   9255
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15002
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.ListBox UserList 
      Height          =   1425
      Left            =   7710
      TabIndex        =   14
      Top             =   360
      Width           =   1650
   End
   Begin VB.ListBox BanList 
      Height          =   2400
      Left            =   7710
      TabIndex        =   13
      Top             =   2235
      Width           =   1650
   End
   Begin VB.CommandButton Command71 
      Caption         =   "&Host"
      Height          =   420
      Left            =   9645
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Text            =   "Chatter"
      Top             =   945
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   75
      TabIndex        =   3
      Top             =   4335
      Width           =   7515
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1635
      Width           =   7395
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1470
      Top             =   15
   End
   Begin VB.ListBox Database 
      Height          =   5325
      Left            =   12015
      TabIndex        =   0
      Top             =   345
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock AServer 
      Left            =   1005
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12340
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   495
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock fReceive 
      Left            =   3870
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock fSend 
      Left            =   3390
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   30000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2460
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   135
      MaxLength       =   20
      TabIndex        =   5
      Top             =   285
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2895
      TabIndex        =   6
      Text            =   "20000"
      Top             =   285
      Width           =   1815
   End
   Begin VB.Image disconnectimg 
      Height          =   330
      Left            =   9450
      Picture         =   "New_Server.frx":5C611
      Top             =   2670
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Image connectimg 
      Height          =   330
      Left            =   9105
      Picture         =   "New_Server.frx":5CC7D
      Top             =   2220
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   4785
      Picture         =   "New_Server.frx":5D234
      Top             =   75
      Width           =   2160
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User List :"
      Height          =   255
      Left            =   7710
      TabIndex        =   16
      Top             =   60
      Width           =   1650
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ban List :"
      Height          =   255
      Left            =   7710
      TabIndex        =   15
      Top             =   1935
      Width           =   1650
   End
   Begin VB.Image exitimg 
      Height          =   330
      Left            =   9540
      Picture         =   "New_Server.frx":5D88E
      Top             =   705
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image exit2img 
      Height          =   330
      Left            =   9525
      Picture         =   "New_Server.frx":5DD3A
      Top             =   1095
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   6600
      MouseIcon       =   "New_Server.frx":5DE05
      Picture         =   "New_Server.frx":5E10F
      Top             =   1000
      Width           =   1080
   End
   Begin VB.Image closeimg 
      Height          =   330
      Left            =   9480
      MouseIcon       =   "New_Server.frx":5E5BB
      Picture         =   "New_Server.frx":5E8C5
      Top             =   1815
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image hostimg 
      Height          =   330
      Left            =   9435
      Picture         =   "New_Server.frx":5EDBD
      Top             =   1455
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image Command7 
      Height          =   330
      Left            =   6555
      Picture         =   "New_Server.frx":5F2D9
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Closed."
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   1020
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Operator Name :"
      Height          =   600
      Left            =   75
      TabIndex        =   7
      Top             =   690
      Width           =   4650
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chat Window :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   75
      TabIndex        =   2
      Top             =   1395
      Width           =   7515
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connection"
      Height          =   495
      Left            =   4860
      TabIndex        =   12
      Top             =   795
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server IP Address :"
      Height          =   570
      Left            =   75
      TabIndex        =   8
      Top             =   60
      Width           =   2610
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server Port  :"
      Height          =   570
      Left            =   2850
      TabIndex        =   9
      Top             =   60
      Width           =   1890
   End
   Begin VB.Menu mnu_UserList 
      Caption         =   "UserList"
      Visible         =   0   'False
      Begin VB.Menu mnu_OP 
         Caption         =   "@"
         Begin VB.Menu mnu_DEOP 
            Caption         =   "DeOp @"
         End
         Begin VB.Menu mnu_OP1 
            Caption         =   "Enable Warning"
         End
         Begin VB.Menu mnu_OP2 
            Caption         =   "Enable Booting"
         End
         Begin VB.Menu mnu_OP3 
            Caption         =   "Enable Both"
         End
      End
      Begin VB.Menu mnu_ServerBan 
         Caption         =   "Ba&n"
         Begin VB.Menu mnu_ServerBanIP 
            Caption         =   "&IP Address"
         End
         Begin VB.Menu mnu_ServerBanUsername 
            Caption         =   "&Username"
         End
      End
      Begin VB.Menu mnu_Boot 
         Caption         =   "Boot"
      End
      Begin VB.Menu mnu_Warn 
         Caption         =   "&Warn"
      End
      Begin VB.Menu mnu_CopyIP 
         Caption         =   "&Copy IP"
      End
   End
   Begin VB.Menu mnu_UnbanTop 
      Caption         =   "UnBan"
      Visible         =   0   'False
      Begin VB.Menu mnu_Unban 
         Caption         =   "&Unban"
      End
   End
   Begin VB.Menu mnu_UserList_ 
      Caption         =   "UserList_"
      Visible         =   0   'False
      Begin VB.Menu mnu_Boot2 
         Caption         =   "&Boot"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Warn2 
         Caption         =   "&Warn"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_SendFile 
         Caption         =   "&Send File..."
      End
   End
   Begin VB.Menu server 
      Caption         =   "Run as &Server"
   End
   Begin VB.Menu space1 
      Caption         =   ""
   End
   Begin VB.Menu client 
      Caption         =   "Run as &Client"
   End
   Begin VB.Menu Space2 
      Caption         =   ""
   End
   Begin VB.Menu help1 
      Caption         =   "Program Info"
      Begin VB.Menu help 
         Caption         =   "Help"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit Program"
      End
   End
End
Attribute VB_Name = "frm_NewServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cc As New frm_SendFile
Dim dd As New frm_SendFile
Dim ips(0 To 50) As Boolean
Dim op As Boolean
Dim pass$, port$, sFilename$, lasttext$
Dim lPos&
Dim buffer() As Byte
Dim inc%, lastsec%

Function CurUserName$()
    Dim sTmp1$
    sTmp1 = Space$(512)
    GetUserName sTmp1, Len(sTmp1)
    CurUserName = Trim$(sTmp1)
End Function

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Private Sub client_Click()
    Picture1.Visible = True
End Sub



Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command7.Picture = closeimg.Picture
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command7.Picture = hostimg.Picture
End Sub

Private Sub Command9__Click()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub AServer_ConnectionRequest(ByVal requestID As Long)
    AServer.Close
    AServer.Accept requestID
End Sub

Private Sub AServer_DataArrival(ByVal bytesTotal As Long)
    Dim data$
    AServer.GetData data$

    If Mid(data$, 1, 5) = "@name" Then
    pos1% = InStr(1, data$, "@password")
    dbname$ = Mid(data$, 6, pos1% - 6)
    dbpw$ = Mid(data$, pos1% + 9)
    datasyn$ = dbname$ & " " & dbpw$
    For i = 0 To Database.ListCount
    If Database.List(i) = datasyn$ Then AServer.SendData "dupe": Exit Sub
    Next i
    Database.AddItem datasyn$
    AServer.SendData "success"
    End If

    If Mid(data$, 1, 5) = "@came" Then
    pos1% = InStr(1, data$, "@password")
    dbname$ = Mid(data$, 6, pos1% - 6)
    dbpw$ = Mid(data$, pos1% + 9)
    datasyn$ = dbname$ & " " & dbpw$
    For i = 0 To Database.ListCount
    If Database.List(i) = datasyn$ Then AServer.SendData "v": Exit Sub
    Next i
    AServer.SendData "nov"
    End If

    If Mid(data$, 1, 5) = "@chng" Then
    pos1% = InStr(1, data$, "@password")
    dbname$ = Mid(data$, 6, pos1% - 6)
    dbpw$ = Mid(data$, pos1% + 9)
    datasyn$ = dbname$ & " " & dbpw$
    For i = 0 To Database.ListCount
    If frm_NewServer.GetListData(Database, True, False, i) = dbname$ Then Database.List(i) = datasyn$: AServer.SendData "@changed": Exit Sub
    Next i
    AServer.SendData "nc"
    End If

End Sub
Private Sub BanList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnu_UnbanTop
End Sub

Private Sub Command1_Click()
    Picture2.Visible = False
End Sub


Private Sub Command7_Click()
    If Command71.Caption = "&Host" Then
    If Winsock3.State <> 0 Then Winsock3.Close
    Winsock3.LocalPort = Text4: Winsock3.Listen
    Command71.Caption = "&Stop"
    Command7.Picture = closeimg.Picture
    Text4.Enabled = False
    pass$ = Text2: port$ = Text4
    AServer.Listen

    Else
    
    Winsock3.Close
    AServer.Close

    For i = 0 To 50
    Winsock2(i).Close
    Next i

    Text4.Enabled = True:
    UserList.Clear
    Command71.Caption = "&Host"
    Command7.Picture = hostimg.Picture
    End If

End Sub
Private Sub Command9_Click()
    Load frm_CreateAccount
    frm_CreateAccount.Show
End Sub

Private Sub exit_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Text2_.PasswordChar = "*"
    frm_NewServer.Height = 5430
    frm_NewServer.Width = 9630
    Picture1.Top = 0
    Picture1.Left = 0
    Picture2.Top = 0
    Picture2.Left = 0
    Timer3.interval = 1000
    Picture1.Visible = True
    StatusBar1.Panels.Item(1).Text = "Click anywhere on the blue background to move window."

    UserList.Clear
    UserList_.Clear

    Me.Show
    Me.SetFocus

    For i = 1 To 50
    Load Winsock2(i)
    ips(i) = False
    Next i

    If Dir(App.Path & "/Database.db") = "" Then GoTo bottom2

    Open App.Path & "/Database.db" For Input As 1
    Do While Not EOF(1)
    Line Input #1, up
    If up <> "" And up <> " " Then Database.AddItem up
    Loop
    Close 1

bottom2:
    Text2 = Winsock3.LocalIP
    lPos = 1
    Text1.Text = CurUserName

End Sub

Private Sub Form_Unload(Cancel As Integer)

    For i = 0 To 50
    Winsock2(i).Close
    If i <> 0 Then Unload Winsock2(i)
    Next i

    Open App.Path & "\Database.db" For Output As 1
    For i = 0 To Database.ListCount
    If Database.List(i) <> "" Then Print #1, Database.List(i)
    Next i
    Close 1

    End

End Sub

Private Sub fSend_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    
    Winsock1.SendData "@prog" & bytesRemaining & "@to" & fSend.RemoteHostIP
    dd.Text2 = "Bytes Remaining : " & bytesRemaining

End Sub

Private Sub help_Click()
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub Image1_Click()
    Unload Me
    End
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = exit2img.Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = exitimg.Picture
End Sub

Private Sub Image4_Click()
Unload Me
End
End Sub

Private Sub IPList_Click()
UserList_.Selected(IPList.ListIndex) = True
End Sub

Private Sub IPList_DblClick()

If IPList.ListIndex = -1 Then Exit Sub
MsgBox IPList.List(IPList.ListIndex), vbInformation

End Sub

Private Sub mnu_Boot_Click()

If UserList.ListIndex = -1 Then Exit Sub

pos1% = InStr(1, UserList.List(UserList.ListIndex), " ")
If pos1% = 0 Then Exit Sub
ctrl = Mid(UserList.List(UserList.ListIndex), 1, pos1% - 1)
bname$ = GetListData(UserList, False, True, UserList.ListIndex)
X$ = InputBox("Enter reason for booting this user.", "Reason?", "Flooded.")

If Len(X$) = 0 Then
Winsock2(ctrl).SendData "@boot"
Else
Winsock2(ctrl).SendData "@mboot" & X$
End If

Pause 0.1

For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@smsg" & " ** Server has booted user " & bname$ & " ** " & vbCrLf: Pause 0.1
Next i

End Sub

Private Sub mnu_Boot2_Click()

If UserList_.ListIndex = -1 Or Winsock1.State <> 7 Or op = False Then Exit Sub
Winsock1.SendData "@userboot" & UserList_.List(UserList_.ListIndex)

End Sub

Private Sub mnu_CopyIP_Click()
Clipboard.SetText IPList.List(IPList.ListIndex)
End Sub

Private Sub mnu_DEOP_Click()

If UserList.ListIndex = -1 Then Exit Sub
ctrl = GetListData(UserList, True, False, UserList.ListIndex)
uname$ = GetListData(UserList, False, True, UserList.ListIndex)
If Mid(uname$, 1, 1) = "@" Then
Winsock2(ctrl).SendData "@OPtake"
Pause 0.1
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@old" & "@" & Mid(uname$, 2) & "@new" & Mid(uname$, 2): Pause 0.1
Next i
UserList.List(UserList.ListIndex) = ctrl & " " & Mid(uname$, 2)
End If

End Sub

Private Sub mnu_OP1_Click()

If UserList.ListIndex = -1 Then Exit Sub
ctrl = GetListData(UserList, True, False, UserList.ListIndex)
uname$ = GetListData(UserList, False, True, UserList.ListIndex)

Dim wasop As Boolean
If Mid(uname$, 1, 1) = "@" Then wasop = True
If wasop <> True Then
uname2$ = "@" & uname$
Else
uname2$ = uname$
End If
Winsock2(ctrl).SendData "@OPgive1"
Pause 0.1
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@old" & uname$ & "@new" & uname2$: Pause 0.1
Next i
UserList.List(UserList.ListIndex) = ctrl & " " & uname2$

End Sub

Private Sub mnu_OP2_Click()

If UserList.ListIndex = -1 Then Exit Sub
ctrl = GetListData(UserList, True, False, UserList.ListIndex)
uname$ = GetListData(UserList, False, True, UserList.ListIndex)

Dim wasop As Boolean
If Mid(uname$, 1, 1) = "@" Then wasop = True
If wasop <> True Then
uname2$ = "@" & uname$
Else
uname2$ = uname$
End If
Winsock2(ctrl).SendData "@OPgive2"
Pause 0.1
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@old" & uname$ & "@new" & uname2$: Pause 0.1
Next i
UserList.List(UserList.ListIndex) = ctrl & " " & uname2$

End Sub

Private Sub mnu_OP3_Click()

If UserList.ListIndex = -1 Then Exit Sub
ctrl = GetListData(UserList, True, False, UserList.ListIndex)
uname$ = GetListData(UserList, False, True, UserList.ListIndex)

Dim wasop As Boolean
If Mid(uname$, 1, 1) = "@" Then wasop = True
If wasop <> True Then
uname2$ = "@" & uname$
Else
uname2$ = uname$
End If
Winsock2(ctrl).SendData "@OPgive3"
Pause 0.1
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@old" & uname$ & "@new" & uname2$: Pause 0.1
Next i
UserList.List(UserList.ListIndex) = ctrl & " " & uname2$

End Sub

Private Sub mnu_OPHelp_Click()

msg$ = "Level 1 - Enable Warning" & vbCrLf
msg$ = msg$ & "Level 2 - Enable Booting" & vbCrLf
msg$ = msg$ & "Level 3 - Enable Warning & Booting" & vbCrLf

MsgBox msg$, vbInformation

End Sub

Private Sub mnu_SendFile_Click()
If IPList.ListIndex = -1 Or IPList.List(IPList.ListIndex) = "" Or Winsock1.State <> 7 Then MsgBox "Please select a valid IP address.", vbInformation: Exit Sub

sFilename$ = OpenDialog(Me, "All Files|*.*", "Locate file to Send...", App.Path)
If Len(sFilename$) = 0 Then Exit Sub
If fSend.State <> 0 Then fSend.Close
fSend.Listen

data2send$ = "@fsip" & IPList.List(IPList.ListIndex) & "@file" & sFilename$
Winsock1.SendData data2send$
End Sub

Private Sub mnu_ServerBanIP_Click()

If UserList.ListIndex = -1 Then Exit Sub

pos1% = InStr(1, UserList.List(UserList.ListIndex), " ")
If pos1% = 0 Then Exit Sub
ctrl = Mid(UserList.List(UserList.ListIndex), 1, pos1% - 1)
IP$ = Winsock2(ctrl).RemoteHostIP
ippos$ = InStrRev(IP$, ".", -1)

For i = 0 To BanList.ListCount
If BanList.List(i) = Mid(IP$, 1, ippos$ - 1) Then Exit Sub
Next i

X$ = InputBox("Enter reason for banning this user.", "Reason?", "Flooded too many times.")

If Len(X$) = 0 Then
Winsock2(ctrl).SendData "@banned"
Else
Winsock2(ctrl).SendData "@mbanned" & X$
End If

BanList.AddItem Mid(IP$, 1, ippos$ - 1)

End Sub

Private Sub mnu_ServerBanUsername_Click()

If UserList.ListIndex = -1 Then Exit Sub

bname$ = GetListData(UserList, False, True, UserList.ListIndex)

For i = 0 To BanList.ListCount
If BanList.List(i) = bname$ Then Exit Sub
Next i

X$ = InputBox("Enter reason for banning this user.", "Reason?", "Flooded too many times.")

If Len(X$) = 0 Then
Winsock2(ctrl).SendData "@banned"
Else
Winsock2(ctrl).SendData "@mbanned" & X$
End If

BanList.AddItem bname$

End Sub


Private Sub mnu_Unban_Click()

If BanList.ListIndex = -1 Then Exit Sub
BanList.RemoveItem BanList.ListIndex

End Sub

Private Sub mnu_Warn_Click()

If UserList.ListIndex = -1 Then Exit Sub

pos1% = InStr(1, UserList.List(UserList.ListIndex), " ")
If pos1% = 0 Then Exit Sub
ctrl = Mid(UserList.List(UserList.ListIndex), 1, pos1% - 1)

X$ = InputBox("Warning Message", App.Title, "This is your 1st warning! Cease your actions immediately.")
Winsock2(ctrl).SendData "@warn" & X$

End Sub

Private Sub mnu_Warn2_Click()

If UserList_.ListIndex = -1 Or Winsock1.State <> 7 Or op = False Then Exit Sub

dat$ = InputBox("Warning Message", App.Title, "This is your 1st warning! Cease your actions immediately.")
If Len(dat$) = 0 Then Exit Sub
Winsock1.SendData "@userwarn" & UserList_.List(UserList_.ListIndex) & "@msg" & dat$

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub server_Click()
Picture1.Visible = False
Picture2.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "&Server" Then
Text6.SetFocus
ElseIf SSTab1.Caption = "&Client" Then
Text6_.SetFocus
End If
End Sub

Private Sub sndfile_Click()
If IPList.ListIndex = -1 Or IPList.List(IPList.ListIndex) = "" Or Winsock1.State <> 7 Then MsgBox "Please select a valid IP address.", vbInformation: Exit Sub

sFilename$ = OpenDialog(Me, "All Files|*.*", "Locate file to Send...", App.Path)
If Len(sFilename$) = 0 Then Exit Sub
If fSend.State <> 0 Then fSend.Close
fSend.Listen

data2send$ = "@fsip" & IPList.List(IPList.ListIndex) & "@file" & sFilename$
Winsock1.SendData data2send$
End Sub

Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Text1__GotFocus()
Text1_.Text = ""
End Sub


Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text2__GotFocus()
Text2_.Text = ""
End Sub

Private Sub Text3__GotFocus()
Text3_.Text = ""
End Sub


Private Sub Text4__Change()

stringcheck$ = "1234567890"

If Len(Text4_) = 0 Then Command7_.Enabled = False: Exit Sub

For i = 1 To Len(Text4_)
If InStr(stringcheck$, Mid(Text4_, i, 1)) = 0 Then Command7_.Enabled = False: Exit Sub
Next i

If Text4_ < 65000 And Text4_ > 0 Then
Command7_.Enabled = True
Else
Command7_.Enabled = False
End If

End Sub

Private Sub Text4_Change()

stringcheck$ = "1234567890"

If Len(Text4) = 0 Then Command7.Enabled = False: Exit Sub

For i = 1 To Len(Text4)
If InStr(stringcheck$, Mid(Text4, i, 1)) = 0 Then Command7.Enabled = False: Exit Sub
Next i

If Text4 < 65000 And Text4 > 0 Then
Command7.Enabled = True
Else
Command7.Enabled = False
End If

End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
KeyAscii = 0

If Mid(Text6, 1, 1) = "/" Then
pos4% = InStr(1, Text6, " ")
If pos4% = 0 Then Text5 = Text5 & " ** Illegal Command **" & vbCrLf: Text5.SelStart = Len(Text5): Exit Sub

afterc$ = LCase(Mid(Text6, 2, pos4% - 2))
If afterc$ <> "msg" Then pos5% = 0: GoTo 9
pos5% = InStr(pos4% + 1, Text6, " ")

9
If pos5% = 0 Then
afterd$ = Mid(Text6, pos4% + 1)
Else
afterd$ = Mid(Text6, pos4% + 1, pos5% - 5)
End If

If afterc$ <> "msg" Then afterd$ = LCase(afterd$)
aftere$ = Mid(Text6, pos5% + 1)

Select Case afterc$
Case "action"
Text5 = Text5 & Text1 & " " & afterd$ & vbCrLf: Text5.SelStart = Len(Text5)
datastring$ = "@action" & Text1 & " " & afterd$ & vbCrLf
Case "msg"
If pos5% = 0 Then Text5 = Text5 & " ** Not enough parameters **" & vbCrLf: Text5.SelStart = Len(Text5): Exit Sub

For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo doover

ctrl = Mid(UserList.List(i), 1, pos1% - 1)
uname$ = Mid(UserList.List(i), pos1% + 1)

If Mid(afterd$, 1, Len(afterd$) - 1) = uname$ Then
Winsock2(ctrl).SendData "@2msg" & "@msg" & aftere$ & "@uname" & Mid(afterd$, 1, Len(afterd$) - 1) & "@fname" & Text1
Text5 = Text5 & " ** Message Sent to " & Mid(afterd$, 1, Len(afterd$) - 1) & " : " & aftere$ & msg$ & " ** " & vbCrLf: Text5.SelStart = Len(Text5)
If inc% > 50 Then Text5.Text = "": inc% = 0
inc% = inc% + 1
GoTo finishmsg
End If

doover:
Next i

Text5 = Text5 & " ** User does not exist **" & vbCrLf:: Text5.SelStart = Len(Text5): Exit Sub

Case Else
Text5 = Text5 & " ** Illegal Command **" & vbCrLf:: Text5.SelStart = Len(Text5): Exit Sub
End Select

If inc% > 50 Then Text5.Text = "": inc% = 0
inc% = inc% + 1

For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData datastring$: Pause 0.1
Next i

finishmsg:
Text6 = ""

Else

Text5 = Text5 & Text1 & " says: " & Text6 & vbCrLf: Text5.SelStart = Len(Text5)

If inc% > 50 Then Text5.Text = "": inc% = 0
inc% = inc% + 1

For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@msg" & Text1 & " says: " & Text6 & vbCrLf: Pause 0.1
Next i

Text6 = ""

End If
End If

End Sub

Private Sub Timer1_Timer()

If AServer.State = 8 Or AServer.State = 9 Then
If Command71.Caption <> "&Host" Then
AServer.Close
AServer.Listen
Else
AServer.Close
End If
End If

For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo checkstat

ctrl = Mid(UserList.List(i), 1, pos1% - 1)
If Winsock2(ctrl).State <> 7 And UserList.List(b) <> "" Then
uname$ = Mid(UserList.List(i), pos1% + 1)

For b = 0 To 50
If Winsock2(b).State = 7 And b <> ctrl And UserList.List(b) <> "" Then Winsock2(b).SendData "@del" & uname$: Pause 0.1
Next b

UserList.RemoveItem i
Winsock2(ctrl).Close
ips(ctrl) = False

End If

DoEvents
Next i

DoEvents

checkstat:

GetState Winsock3, Label6

For i = 0 To 50
If Winsock2(i).State <> 7 Then Winsock2(i).Close
Next i

End Sub

Private Sub Timer2_Timer()

If Winsock1.State = 9 Or Winsock1.State = 8 Then SetValues False

If fSend.State = 8 Then fSend.Close: Close 1, 2
If fReceive.State = 8 Then fReceive.Close: Close 1, 2: lPos = 1: cc.Command1.Enabled = True: cc.Text2 = "Done."

GetState Winsock1, Label6_

End Sub

Private Sub Timer3_Timer()
StatusBar1.Panels.Item(3).Text = Time
If frm_NewServer.Width = 1 Then
Else
frm_NewServer.Caption = ""
End If
If Picture1.Visible = True Then
StatusBar1.Panels.Item(2).Text = Label6_.Caption
Else
StatusBar1.Panels.Item(2).Text = Label6.Caption
End If
End Sub

Private Sub UserList__Click()

If UserList_.ListIndex = -1 Then Exit Sub
IPList.Selected(UserList_.ListIndex) = True

End Sub

Private Sub UserList__MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then PopupMenu mnu_UserList_

End Sub

Private Sub UserList_Click()

If UserList.ListIndex <> -1 Then
pos1% = InStr(1, UserList.List(UserList.ListIndex), " ")
If pos1% = 0 Then Exit Sub

ctrl = Mid(UserList.List(UserList.ListIndex), 1, pos1% - 1)
End If

End Sub

Private Sub UserList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then PopupMenu mnu_UserList

End Sub
Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)

For i = 0 To 50
If Winsock2(i).State = 0 Then
Winsock2(i).Accept requestID
ips(i) = False
GoTo ending
End If
DoEvents
Next i
ending:

End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data$
Winsock2(Index).GetData data$

''''''''''''''''''''''''''''''''''''''''''
If Mid(data$, 1, 5) = "@pass" Then
IP$ = Winsock2(Index).RemoteHostIP
ippos$ = InStrRev(IP$, ".", -1)
pos1% = InStr(1, data$, "@name")
inname$ = Mid(data$, 6, Len(data$) - Len(Mid(data$, pos1%)) - 5)
inpass$ = Mid(data$, pos1% + 5)
For i = 0 To BanList.ListCount
If BanList.List(i) = Mid(IP$, 1, ippos$ - 1) Or BanList.List(i) = inname$ Then Winsock2(Index).SendData "@ubanned": Exit Sub
Next i
For z = 0 To Database.ListCount
outname$ = GetListData(Database, True, False, z)
outpw$ = GetListData(Database, False, True, z)
If inpass$ = outpw$ And inname$ = outname$ Then
For i = 0 To UserList.ListCount - 1
currname$ = GetListData(UserList, False, True, i)
If Mid(currname$, 1, 1) = "@" Then currname$ = Mid(currname$, 2)
If currname$ = inname$ Then
Winsock2(Index).SendData "@nameused"
Do Until Winsock2(Index).State <> 7
Pause 0.2
Loop
Winsock2(Index).Close: Exit Sub
End If
Next i
UserList.AddItem Index & " " & inname$
For i = 0 To 50
If i = Index Then GoTo skipit
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@newuser" & inname$: Pause 0.1
skipit:
Next i
For i = 0 To UserList.ListCount - 1
pos1% = InStr(1, UserList.List(i), " ")
users$ = users$ & "@user" & Mid(UserList.List(i), pos1% + 1)
Next i
ips(Index) = True
Winsock2(Index).SendData users$
GoTo verified
End If
Next z
Winsock2(Index).SendData "@outpw": Exit Sub
End If


If ips(Index) = True Then
GoTo verified
Else
Winsock2(Index).Close: Exit Sub
End If




verified:
If Mid(data$, 1, 4) = "@old" Then
pos9% = InStr(1, data$, "@new")
If pos9% = 0 Then Exit Sub
oldname$ = Mid(data$, 5, pos9% - 5)
newname$ = Mid(data$, pos9% + 4)
For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo doover3
pname$ = Mid(UserList.List(i), 1, pos1%)
uname$ = Mid(UserList.List(i), pos1% + 1)
If uname$ = oldname$ Then UserList.List(i) = pname$ & newname$: GoTo next2
doover3:
Next i
next2:
For d = 0 To 50
If Winsock2(d).State = 7 Then Winsock2(d).SendData data$: Pause 0.1
Next d
End If


If Mid(data$, 1, 4) = "@msg" Then
Text5 = Text5 & Mid(data$, 5): Text5.SelStart = Len(Text5)
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData "@msg" & Mid(data$, 5): Pause 0.1
Next i
End If


If Mid(data$, 1, 7) = "@action" Then
If inc% > 50 Then Text5.Text = "": inc% = 0
Text5 = Text5 & Mid(data$, 8): Text5.SelStart = Len(Text5): inc% = inc% + 1
For i = 0 To 50
If Winsock2(i).State = 7 Then Winsock2(i).SendData data$: Pause 0.1
Next i
End If

If Mid(data$, 1, 5) = "@2msg" Then
pos6% = InStr(5, data$, "@msg")
If pos6% = 0 Then Exit Sub
pos7% = InStr(5, data$, "@uname")
If pos7% = 0 Then Exit Sub
pos8% = InStr(5, data$, "@fname")
If pos8% = 0 Then Exit Sub
tmsg$ = Mid(data$, pos6% + 4, pos7% - 10)
tuname$ = Mid(data$, pos7% + 6, pos8% - pos7% - 6)
fname$ = Mid(data, pos8% + 6)
For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo doover
ctrl = Mid(UserList.List(i), 1, pos1% - 1)
uname$ = Mid(UserList.List(i), pos1% + 1)
If fname$ = uname$ Then
Winsock2(ctrl).SendData "@2msg" & "@msg" & tmsg$ & "@uname" & tuname$ & "@fname" & fname$
If inc% > 50 Then Text5.Text = "": inc% = 0
inc% = inc% + 1
End If
doover:
Next i
End If

''''''''''''''''''''''''''''''''''''''''''
If Mid(data$, 1, 8) = "@fdenied" Then
For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo nextone2
ctrl = Mid(UserList.List(i), 1, pos1% - 1)
If Mid(data$, 9) = Winsock2(ctrl).RemoteHostIP Then
Winsock2(ctrl).SendData "@fdenied"
End If
nextone2:
DoEvents
Next i
End If

If Mid(data$, 1, 5) = "@fsip" Then
fpos1% = InStr(1, data$, "@file"): If fpos1% = 0 Then Exit Sub
fsip$ = Mid(data$, 6, fpos1% - 6)
fsfile$ = Mid(data$, fpos1% + 5)
For i = 0 To UserList.ListCount
If GetListData(UserList, True, False, i) = Index Then sname$ = GetListData(UserList, False, True, i): GoTo skip2:
Next i
skip2:
For i = 0 To UserList.ListCount
ctrl = GetListData(UserList, True, False, i)
If fsip$ = Winsock2(ctrl).RemoteHostIP And Winsock2(ctrl).State = 7 Then
Winsock2(ctrl).SendData "@fip" & Winsock2(Index).RemoteHostIP & "@file" & fsfile$ & "@from" & sname$: Exit Sub
End If
nextone:
DoEvents
Next i
End If

If Mid(data$, 1, 5) = "@prog" Then
pos1% = InStr(1, data$, "@to"): If pos1% = 0 Then Exit Sub
brem$ = Mid(data$, 6, pos1% - 6)
toip$ = Mid(data$, pos1% + 3)
For i = 0 To UserList.ListCount
ctrl = GetListData(UserList, True, False, i)
If Winsock2(ctrl).RemoteHostIP = toip$ Then Winsock2(Index).SendData "@fprog" & brem$
Next i
End If




'''' User @ Commands
If Mid(data$, 1, 6) = "@getip" Then
For v% = 0 To UserList.ListCount
If GetListData(UserList, False, True, v%) = Mid(data$, 7) Then Winsock2(Index).SendData "@ipuser" & Winsock2(GetListData(UserList, True, False, v%)).RemoteHostIP & "@username" & Mid(data$, 7)
DoEvents
Next v%
End If

If Mid(data$, 1, 9) = "@userboot" Then
For i = 0 To UserList.ListCount
pos0% = InStr(1, UserList.List(i), " ")
If pos0% = 0 Then GoTo nextone6
ctrl = Mid(UserList.List(i), 1, pos0% - 1)
usename$ = Mid(UserList.List(i), pos0% + 1)
If ctrl = Index Then tusername$ = usename$
nextone6:
DoEvents
Next i
For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo nextone4
ctrl = Mid(UserList.List(i), 1, pos1% - 1)
If Mid(data$, 10) = Mid(UserList.List(i), pos1% + 1) Then
Winsock2(ctrl).SendData "@boot"
For c = 0 To 50
If Winsock2(c).State = 7 Then Winsock2(c).SendData "@smsg" & " ** " & tusername$ & " has booted user " & Mid(data$, 10) & " ** " & vbCrLf: Pause 0.1
Next c
Exit Sub
End If
nextone4:
DoEvents
Next i
End If

If Mid(data$, 1, 9) = "@userwarn" Then
pos2% = InStr(1, data$, "@msg")
If pos2% = 0 Then Exit Sub
usernam$ = Mid(data$, 10, pos2% - 10)
messag$ = Mid(data$, pos2% + 4)
For i = 0 To UserList.ListCount
pos1% = InStr(1, UserList.List(i), " ")
If pos1% = 0 Then GoTo nextone5
ctrl = Mid(UserList.List(i), 1, pos1% - 1)
If usernam$ = Mid(UserList.List(UserList.ListIndex), pos1% + 1) Then Winsock2(ctrl).SendData "@warn" & messag$
nextone5:
DoEvents
Next i
End If


''''''''''' Tic - Tac - Toe
If Mid(data$, 1, 7) = "@reqttt" Then
pos13% = InStr(8, data$, "@from"): If pos13% = 0 Then Exit Sub
toip$ = Mid(data$, 8, pos13% - 8)
fromname$ = Mid(data$, pos13% + 5)
For i = 0 To UserList.ListCount
If Winsock2(GetListData(UserList, True, False, i)).RemoteHostIP = toip$ Then
Winsock2(GetListData(UserList, True, False, i)).SendData "@reqttt" & fromname$
Exit Sub
End If
Next i
End If

If Mid(data$, 1, 7) = "@accttt" Then
For i = 0 To UserList.ListCount
If GetListData(UserList, False, True, i) = Mid(data$, 8) Then Winsock2(GetListData(UserList, True, False, i)).SendData "@accttt" & Winsock2(Index).RemoteHostIP:  Exit Sub
Next i
End If

If Mid(data$, 1, 7) = "@denttt" Then
For i = 0 To UserList.ListCount
If GetListData(UserList, False, True, i) = Mid(data$, 8) Then Winsock2(GetListData(UserList, True, False, i)).SendData "@denttt": Exit Sub
Next i
End If

If Mid(data$, 1, 9) = "@tttclick" Then
pos1% = InStr(1, data$, "@spot")
playerip$ = Mid(data$, 10, pos1% - 10)
spotclick$ = Mid(data, pos1% + 5)
For i = 0 To UserList.ListCount
ctrl = GetListData(UserList, True, False, i)
If Winsock2(ctrl).RemoteHostIP = playerip$ Then Winsock2(ctrl).SendData "@tttclick" & spotclick$:  Exit Sub
Next i
End If

If Mid(data$, 1, 7) = "@tttend" Then
For i = 0 To UserList.ListCount
ctrl = GetListData(UserList, True, False, i)
If Winsock2(ctrl).RemoteHostIP = Mid(data$, 8) Then Winsock2(ctrl).SendData "@tttend": Exit Sub
Next i
End If

If Mid(data$, 1, 7) = "@tttwin" Then
pos1% = InStr(1, data$, "@who")
userip$ = Mid(data$, 8, pos1% - 7)
whowon$ = Mid(data$, pos1% + 4)
For i = 0 To UserList.ListCount
ctrl = GetListData(UserList, True, False, i)
If Winsock2(ctrl).RemoteHostIP = userip$ Then Winsock2(ctrl).SendData "@tttwin" & whowon$: Exit Sub
Next i
End If

End Sub
Private Sub Command7__Click()

If Command7_.Picture = connectimg.Picture Then
If Winsock1.State <> 0 Then Winsock1.Close
Winsock1.Connect Text3_, Text4_
Command7_.Picture = disconnectimg.Picture
TextNable False
Else
SetValues False
End If

End Sub

Sub TextNable(choice As Boolean)
Text1_.Enabled = choice
Text2_.Enabled = choice
Text3_.Enabled = choice
Text4_.Enabled = choice
End Sub

Private Sub fReceive_DataArrival(ByVal bytesTotal As Long)

Dim buffer2() As Byte
fReceive.GetData buffer2()

Put #2, lPos, buffer2()
lPos = lPos + UBound(buffer2) + 1

End Sub


Private Sub fSend_ConnectionRequest(ByVal requestID As Long)

fSend.Close
fSend.Accept requestID

dd.Show
dd.Caption = "Sending File..."
dd.Text1 = sFilename$

Open sFilename$ For Binary Access Read As 1
ReDim buffer(LOF(1))
Get #1, 1, buffer()
Close 1

fSend.SendData buffer()

End Sub
Private Sub fSend_SendComplete()

fSend.Close
ReDim buffer(0)
Text5_ = Text5_ & " ** Send File : File Sent **" & vbCrLf: Text5_.SelStart = Len(Text5_)

dd.Command1.Enabled = True
dd.Text2 = "Done."

End Sub

Private Sub Text1__Change()
txt$ = Text1_
Text1_ = StringChange(txt$)
End Sub

Private Sub Text6__KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
KeyAscii = 0

If Winsock1.State <> 7 Then Text6_ = "": Exit Sub
If lasttext$ = Text6_ Then MsgBox "Why are you saying the same thing?", vbQuestion: Text6_ = "": Exit Sub

lasttext$ = Text6_

If Mid(Text6_, 1, 1) = "/" Then
pos4% = InStr(1, Text6_, " ")
If pos4% = 0 Then Text5_ = Text5_ & " ** Illegal Command **" & vbCrLf: Text5_.SelStart = Len(Text5_): Exit Sub

afterc$ = LCase(Mid(Text6_, 2, pos4% - 2))
If afterc$ <> "msg" Then pos5% = 0: GoTo 9
pos5% = InStr(pos4% + 1, Text6_, " ")

9
If pos5% = 0 Then
afterd$ = Mid(Text6_, pos4% + 1)
Else
afterd$ = Mid(Text6_, pos4% + 1, pos5% - 5)
End If

If afterc$ <> "msg" And afterc$ <> "change" Then afterd$ = LCase(afterd$)
aftere$ = Mid(Text6_, pos5% + 1)

Select Case afterc$
Case "action"
datastring$ = "@action" & Text1_ & " " & afterd$ & vbCrLf
Case "msg"
If pos5% = 0 Then Text5_ = Text5_ & " ** Not enough parameters **" & vbCrLf: Text5_.SelStart = Len(Text5_): Exit Sub

For i = 0 To UserList_.ListCount
uname$ = UserList_.List(i)

If Mid(afterd$, 1, Len(afterd$) - 1) = uname$ Then
Winsock1.SendData "@2msg" & "@msg" & aftere$ & "@uname" & Text1_ & "@fname" & Mid(afterd$, 1, Len(afterd$) - 1)
Text5_ = Text5_ & " ** Message Sent to " & Mid(afterd$, 1, Len(afterd$) - 1) & " : " & aftere$ & msg$ & " ** " & vbCrLf: Text5_.SelStart = Len(Text5_)
If inc% > 50 Then Text5_.Text = "": inc% = 0
inc% = inc% + 1
GoTo finishmsg
End If

doover:
Next i

Text5_ = Text5_ & " ** User does not exist **" & vbCrLf: Text5_.SelStart = Len(Text5_): Exit Sub

Case "change"
oldname$ = Text1_
newname$ = StringChange(afterd$)
If Mid(newname$, 1, 1) = "@" Then newname$ = Mid(newname$, 2)

For i = 0 To UserList_.ListCount
currname2$ = UserList_.List(i)
If Mid(currname2$, 1, 1) = "@" Then currname2$ = Mid(currname2$, 2)
If currname2$ = newname$ Or currname1$ = newname$ Then Text5_ = Text5_ & " ** User already exists in list **" & vbCrLf: Text5_.SelStart = Len(Text5_): Exit Sub
Next i

If op = True Then newname$ = "@" & newname$

For i = 0 To UserList_.ListCount
If UserList_.List(i) = oldname$ Then UserList_.List(i) = newname$: Text1_ = newname$
Next i

Winsock1.SendData "@old" & oldname$ & "@new" & newname$: Text6_ = "": Exit Sub

Case Else
Text5_ = Text5_ & " ** Illegal Command **" & vbCrLf:: Text5_.SelStart = Len(Text5_): Exit Sub
End Select

If inc% > 50 Then Text5_.Text = "": inc% = 0
inc% = inc% + 1

If Winsock1.State = 7 Then Winsock1.SendData datastring$


finishmsg:
Text6_ = ""

Else

If inc% > 50 Then Text5_.Text = "": inc% = 0
inc% = inc% + 1

If Winsock1.State = 7 Then Winsock1.SendData "@msg" & Text1_ & " says: " & Text6_ & vbCrLf
Text6_ = ""

End If

Text6_.Enabled = False
Text6_ = "Sending Message..."
Pause 1.2
Text6_.Enabled = True
Text6_ = ""
Text6_.SetFocus

End If


End Sub
Private Sub winsock1_Connect()

data$ = "@pass" & Text1_ & "@name" & Text2_
Winsock1.SendData data$

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim data$
Winsock1.GetData data$

'' Server Responses
If data$ = "@outpw" Then Winsock1.Close: MsgBox "Incorrect Username/Password.", vbInformation: SetValues False
If InStr(1, data$, "@nameused") > 0 Then Winsock1.Close: MsgBox "Name currently in use.", vbInformation: SetValues False
If data$ = "@ubanned" Then Winsock1.Close: MsgBox "Cannot connect to server. You have been banned.", vbExclamation: SetValues False


'' Messages
If Mid(data$, 1, 4) = "@msg" Then
If inc% > 50 Then Text5_.Text = "": inc% = 0
Text5_ = Text5_ & Mid(data$, 5): Text5_.SelStart = Len(Text5_): inc% = inc% + 1
End If

If Mid(data$, 1, 7) = "@action" Then
If inc% > 50 Then Text5_.Text = "": inc% = 0
Text5_ = Text5_ & Mid(data$, 8): Text5_.SelStart = Len(Text5_): inc% = inc% + 1
End If

If Mid(data$, 1, 5) = "@2msg" Then
pos6% = InStr(5, data$, "@msg")
If pos6% = 0 Then Exit Sub
pos7% = InStr(5, data$, "@uname")
If pos7% = 0 Then Exit Sub
pos8% = InStr(5, data$, "@fname")
If pos8% = 0 Then Exit Sub
tmsg$ = Mid(data$, pos6% + 4, pos7% - 10)
tuname$ = Mid(data$, pos7% + 6, pos8% - pos7% - 6)
fname$ = Mid(data, pos8% + 6)
If inc% > 50 Then Text5_.Text = "": inc% = 0
inc% = inc% + 1
Text5_ = Text5_ & " ** Message from " & tuname$ & " : " & tmsg$ & " ** " & vbCrLf: Text5_.SelStart = Len(Text5_): inc% = inc% + 1
End If

If Mid(data$, 1, 5) = "@smsg" Then
If inc% > 50 Then Text5_.Text = "": inc% = 0
inc% = inc% + 1
Text5_ = Text5_ & Mid(data$, 6): Text5_.SelStart = Len(Text5_): inc% = inc% + 1
End If


'' Server Commands
If Mid(data$, 1, 5) = "@boot" Then Winsock1.Close: MsgBox "You were booted.", vbExclamation: SetValues False
If Mid(data$, 1, 6) = "@mboot" Then Winsock1.Close: MsgBox "You were booted." & Chr(13) & Chr(13) & "Reason : " & Mid(data$, 7), vbExclamation: SetValues False
If Mid(data$, 1, 5) = "@warn" Then MsgBox Mid(data$, 6), vbExclamation
If Mid(data$, 1, 8) = "@mbanned" Then Winsock1.Close: MsgBox "Server has banned you." & Chr(13) & Chr(13) & "Reason : " & Mid(data$, 9), vbExclamation: SetValues False
If data$ = "@banned" Then Winsock1.Close: MsgBox "Server has banned you.", vbExclamation: SetValues False

If Mid(data$, 1, 7) = "@OPgive" Then
Level$ = Mid(data, 8, 1)
If Level$ = "1" Then
mnu_Warn2.Enabled = True: mnu_Boot2.Enabled = False: op = True: Text1_ = "@" & Text1_
ElseIf Level$ = "2" Then
mnu_Boot2.Enabled = True: mnu_Warn2.Enabled = False: op = True: Text1_ = "@" & Text1_
ElseIf Level$ = "3" Then
mnu_Boot2.Enabled = True: mnu_Warn2.Enabled = True:  op = True: Text1_ = "@" & Text1_
End If
End If

If data$ = "@OPtake" Then mnu_Warn2.Enabled = False: mnu_Warn2.Enabled = False:  op = False: Text1_ = Mid(Text1_, 2)
If Mid(data$, 1, 7) = "@ipuser" Then
pos3% = InStr(7, data$, "@username")
usename$ = Mid(data$, pos3% + 9)
useip$ = Mid(data$, 8, pos3% - 8)
For q = 0 To UserList_.ListCount
If UserList_.List(q) = usename$ Then IPList.List(q) = useip$
Next q
End If


'' Server User Responses
If Mid(data$, 1, 4) = "@del" Then
For i = 0 To UserList_.ListCount
If UserList_.List(i) = Mid(data$, 5) Then If Len(data$) > 5 Then UserList_.RemoveItem i: IPList.RemoveItem i
Next i
End If

If Mid(data$, 1, 8) = "@newuser" Then UserList_.AddItem Mid(data$, 9):  Winsock1.SendData "@getip" & Mid(data$, 9): Pause 0.2
If Mid(data$, 1, 5) = "@user" Then ExtractUsers data$
If Mid(data$, 1, 4) = "@old" Then
pos9% = InStr(1, data$, "@new")
If pos9% = 0 Then Exit Sub
oldname$ = Mid(data$, 5, pos9% - 5)
newname$ = Mid(data$, pos9% + 4)
For i = 0 To UserList_.ListCount
If UserList_.List(i) = oldname$ Then UserList_.List(i) = newname$ ': Text1 = newname$
Next i
End If


'' File Sending
redodis:
If Mid(data$, 1, 6) = "@fprog" Then
nextpos2% = InStr(7, data$, "@fprog")
If nextpos2% = 0 Then
cc.Text2 = "Bytes Remaining : " & Mid(data$, 7)
Else
cc.Text2 = "Bytes Remaining : " & Mid(data$, 7, nextpos2% - 6)
data$ = Mid(data$, nextpos2%): GoTo redodis
End If
End If

If Mid(data$, 1, 8) = "@fdenied" Then MsgBox "User has denied your file request.", vbCritical: Close 1, 2

If Mid(data$, 1, 4) = "@fip" Then
fpos1% = InStr(1, data$, "@file"): If fpos1% = 0 Then Exit Sub
fpos2% = InStr(1, data$, "@from"): If fpos2% = 0 Then Exit Sub
IPNumb$ = Mid(data$, 5, fpos1% - 5)
fname$ = Mid(data$, fpos1% + 5, fpos2% - 9)
fromname$ = Mid(data$, fpos2% + 6)
lRet% = MsgBox("Accept incoming file from " & fromname$ & "?" & vbCrLf & fname$, vbYesNo + vbQuestion, "Server")
If lRet = vbYes Then
If fReceive.State <> 0 Then fReceive.Close
pos1% = InStrRev(fname$, "\", -1)
If pos1% = 0 Then Exit Sub
fname$ = SaveDialog(Me, "All Files|*.*", "Save Incoming File As...", "")
If Len(fname$) = 0 Then GoTo filed
cc.Show
cc.Caption = "Getting File..."
cc.Text1 = fname$
Open fname$ For Binary Access Write As 2
fReceive.Connect IPNumb$, "30000"
Else
filed:
Winsock1.SendData "@fdenied" & IPNumb$
Close 1, 2
End If
End If


'''''''' Tic - Tac - Toe
If Mid(data$, 1, 7) = "@reqttt" Then
lRet = MsgBox("Access Tic-Tac-Toe request from " & Mid(data$, 8) & "?", vbInformation + vbYesNo)
If lRet = vbYes Then
For i = 0 To UserList_.ListCount
If UserList_.List(i) = Mid(data$, 8) Then ttt1 = IPList.List(i)
Next i
ttt2 = Winsock1.LocalIP
player2TTT = True
played = True
tttPlayer2.Show
tttPlayer2.Frame11.Caption = Text1_
tttPlayer2.Frame12.Caption = Mid(data$, 8)
Winsock1.SendData "@accttt" & Mid(data$, 8)
Else
Winsock1.SendData "@denttt" & Mid(data$, 8)
End If
End If

If Mid(data$, 1, 7) = "@accttt" Then
ttt1 = Winsock1.LocalIP
ttt2 = Mid(data$, 8)
For i = 0 To IPList.ListCount
If IPList.List(i) = Mid(data$, 8) Then p2$ = UserList_.List(i)
Next i
If p2$ = "" Then MsgBox "Error. Player 2 was not located.", vbCritical: Exit Sub
player1TTT = True
played = False
tttPlayer1.Show
tttPlayer1.Frame12.Caption = Text1_
tttPlayer1.Frame11.Caption = p2$
End If

If Mid(data$, 1, 7) = "@denttt" Then MsgBox "Your request for Tic-Tac-Toe was denied.", vbCritical

If Mid(data$, 1, 9) = "@tttclick" Then
If player1TTT = True Then
tttPlayer1.Spot(Mid(data$, 10)) = "O"
spots(Mid(data$, 10)) = False
played = False
If player1TTT = False Then player1TTT = True
ElseIf player2TTT = True Then
tttPlayer2.Spot(Mid(data$, 10)) = "X"
spots(Mid(data$, 10)) = False
If player2TTT = False Then player2TTT = True
played = False
End If
End If

If data$ = "@tttend" Then Unload tttPlayer1: Unload tttPlayer2: MsgBox "User has closed Tic-Tac-Toe on other side.", vbInformation

End Sub

Sub ExtractUsers(users$)

nextu% = InStr(1, users$, "@user")
nextu2% = InStr(nextu% + 1, users$, "@user")

Do Until nextu% = 0 Or numbspc% > Len(users$)

'(41) = "@userCoozzzzz (14) @userCoozzzzz2 (29) @userCoozzzzz3"

If nextu2% = 0 Then
UserList_.AddItem Mid(users$, nextu% + 5, Len(users$)): Pause 0.1
Winsock1.SendData "@getip" & Mid(users$, nextu% + 5, Len(users$))
Else
UserList_.AddItem Mid(users$, nextu% + 5, nextu2% - nextu% - 5): Pause 0.1
Winsock1.SendData "@getip" & Mid(users$, nextu% + 5, nextu2% - nextu% - 5)
End If

old2% = nextu2%
nextu% = InStr(nextu% + 6, users$, "@user")
nextu2% = InStr(nextu% + 8, users$, "@user")
DoEvents

Loop

End Sub

Function StringChange(str$)

redo:
checkstring$ = "!#$%^&*()-=][{}\|/`~1234567890QABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwyxz"

For i = 1 To Len(str$)
If InStr(1, checkstring, Mid(str$, i, 1)) = 0 Then
str$ = Mid(str$, 1, i - 1) & Mid(str$, i + 1)
GoTo redo
End If
Next i

If op = True Then str$ = "@" & str$
StringChange = str$


End Function

Function GetListData(uname As ListBox, ctrl As Boolean, username As Boolean, i)

If i = -1 Then Exit Function
pos1% = InStr(1, uname.List(i), " ")
If pos1% = 0 Then Exit Function

ctrl1 = Mid(uname.List(i), 1, pos1% - 1)
lname = Mid(uname.List(i), pos1% + 1)

If ctrl = True Then GetListData = ctrl1
If username = True Then GetListData = lname

End Function

Sub SetValues(Connected As Boolean)

If Connected = False Then

Winsock1.Close
UserList_.Clear
IPList.Clear
Command7_.Picture = connectimg.Picture
op = False
TextNable True
Text5_ = ""

End If

End Sub

Sub GetState(ws As Winsock, thelab As Label)


Select Case ws.State
Case 0
thelab.Caption = "Closed."
Case 1
thelab.Caption = "Listening."
Case 2
thelab.Caption = "Open."
Case 3
thelab.Caption = "Connection Pending."
Case 4
thelab.Caption = "Resolving Host."
Case 5
thelab.Caption = "Host Resolved."
Case 6
thelab.Caption = "Connecting."
Case 7
thelab.Caption = "Connected."
Case 8
thelab.Caption = "Closing."
Case 9
thelab.Caption = "Error."
DoEvents
End Select

End Sub
