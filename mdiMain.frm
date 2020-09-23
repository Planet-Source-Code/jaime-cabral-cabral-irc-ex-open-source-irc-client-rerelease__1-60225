VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Cabral IRC"
   ClientHeight    =   10920
   ClientLeft      =   3615
   ClientTop       =   2355
   ClientWidth     =   13080
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picVBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8715
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   8715
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   330
      Width           =   60
   End
   Begin MSComctlLib.ImageList imgView 
      Left            =   4440
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":16C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17006
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":173A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1773A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2CC46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2DEC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2F14A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8715
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   2280
      TabIndex        =   4
      Top             =   330
      Width           =   2340
      Begin MSComctlLib.TreeView tvNotify 
         Height          =   4695
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   8281
         _Version        =   393217
         Style           =   7
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   50
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   3600
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvMain 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   6165
         _Version        =   393217
         Indentation     =   88
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imgView"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer timeParse 
      Left            =   5520
      Top             =   2640
   End
   Begin MSComctlLib.ImageList imgTaskbar 
      Left            =   4440
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   12
      MaskColor       =   16711935
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2F4E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2F748
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2FAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2FD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":300E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   9045
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer ISON 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5520
      Top             =   2160
   End
   Begin VB.PictureBox picRaw 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   13050
      TabIndex        =   1
      Top             =   9345
      Visible         =   0   'False
      Width           =   13080
      Begin VB.TextBox txtRaw 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   6015
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Options"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Channel Folder"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Script"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "List Channels"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "URL Catcher"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send File"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Direct Chat"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tile Windows"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cascade Windows"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help!"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         Left            =   4560
         TabIndex        =   9
         Text            =   "<search the web>"
         Top             =   0
         Width           =   5055
      End
   End
   Begin MSWinsockLib.Winsock CHATx 
      Index           =   0
      Left            =   5040
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock CHAT 
      Index           =   0
      Left            =   5040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcp 
      Left            =   5040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ident 
      Index           =   0
      Left            =   5040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3067A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":316CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32722
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":33EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":34F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":35F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":36FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":37FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3AD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3CA16
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3E722
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4142E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":444B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":45A46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCabral 
      Caption         =   "&Cabral"
      Begin VB.Menu mnuGeneralOptions 
         Caption         =   "General Options"
      End
      Begin VB.Menu h0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Web Page"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTile 
         Caption         =   "Tile"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWList 
         Caption         =   "List"
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRaw 
         Caption         =   "Raw Window"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim elength As Double
Dim el As Long, b1 As Long, b2 As Long, b3 As Long, b4 As Long


Dim bDrag As Boolean    'True if the user has the mouse pressed while on the resize bars


Private Sub CHAT_Close(Index As Integer)
    'MsgBox Index
    Unload CHAT(Index)
    ChatWindowName(Index) = ""
    'Call DoColor(ChatWindow(Index).txtDCC, "4* Connection terminated")
    Call DoColor(ChatWindow(Index).txtDCC, "4* DCC session closed!")
End Sub

Private Sub CHAT_Connect(Index As Integer)
    'Call DoColor(ChatWindow(Index).txtDCC, "4* Connection established")
    Call DoColor(ChatWindow(Index).txtDCC, "* DCC session established!")
    'DoColor frmStatus.txtStatus, "CHAT CONNECT"
End Sub

Private Sub CHAT_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    CHAT(Index).GetData strData
    'ChatWindow(Index).txtDCC.SelText = "<" & ChatWindowName(Index) & "> " & strData
    strData = Replace(strData, Chr(10), "")
    strData = Replace(strData, Chr(13), "")
    DoColor ChatWindow(Index).txtDCC, "0<" & nickname & "> " & strData
End Sub

Private Sub CHATx_Close(Index As Integer)
    Call DoColor(ChatWindowx(Index).txtDCC, "4* DCC session closed!")
    CHATx(Index).Close
    ChatWindowNamex(Index) = ""
End Sub

Private Sub CHATx_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'MsgBox requestID
    CHATx(Index).Close
    CHATx(Index).Accept requestID
    Call DoColor(ChatWindowx(Index).txtDCC, "4* DCC session established!")
End Sub

Private Sub CHATx_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    CHATx(Index).GetData strData
    strData = Replace(strData, Chr(10), "")
    'ChatWindowx(Index).txtDCC.SelText = "<" & ChatWindowNamex(Index) & "> " & strData & vbCrLf
    DoColor ChatWindowx(Index).txtDCC, "0<" & ChatWindowNamex(Index) & "> " & strData
End Sub

Private Sub cToolbar1_ButtonClick(ByVal lButton As Long)
   'MsgBox "Toolbar1 ButtonClick:" & vbTab & CStr(lButton) & ",Pressed=" & cToolbar1.ButtonPressed(lButton) & ",Checked=" & cToolbar1.ButtonChecked(lButton)
   Dim xServer As String
   Dim xPort As Integer
   Dim i As Integer
   
   'load previous nickname/server/real name
    server = ReadINI("INFO", "SERVER")
    nickname = ReadINI("INFO", "NICKNAME")
    email = ReadINI("INFO", "USERNAME")
    RealName = ReadINI("INFO", "REALNAME")
   
   
   If (lButton = 0) Then
        If tcp.State <> 0 Then
            tcp.Close
        End If
        For i = 1 To Len(server)
            If Mid(server, i, 1) = ":" Then
                PortFound = True
                xServer = Mid(server, 1, i - 1)
                xPort = Val(Mid(server, i + 1))
            End If
        Next i
        If PortFound Then
            Call DoColor(frmStatus.txtStatus, "2Attempting to connect to4 " & xServer & " 1(10" & xPort & "1)")
            tcp.Connect xServer, xPort
        Else
            Call DoColor(frmStatus.txtStatus, "2Attempting to connect to4 " & server & " 1(1066671)")
            tcp.Connect server, 6667
        End If
   End If
    If (lButton = 4) Then
        frmChannelFolder.SHOW 0, mdiMain
    End If
    If (lButton = 10) Then
        Call mnuTile_Click
    End If
    If (lButton = 11) Then
        Call mnuCascade_Click
    End If
End Sub




Private Sub ident_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    
    Dim socket As Variant
    For Each socket In ident
        If (socket.State = sckClosed Or socket.State = sckError) Then
            socket.Close
            socket.Accept requestID
            'frmOptions.txtIdentLog = frmOptions.txtIdentLog & socket.RemoteHostIP & "[" & requestID & "]" & vbCrLf
            socket.SendData socket.LocalPort & ", " & requestID & ":USERID:WIN32:" & IdentUserID & vbCrLf
            '1236, 7000 : USERID : UNIX : higher
            Call DoColor(frmStatus.txtStatus, "6* Ident request from " & mdiMain.ident(0).RemoteHostIP)
            Call DoColor(frmStatus.txtStatus, "6* Ident Reply: " & socket.LocalPort & ", " & requestID & ":USERID:WIN32:" & IdentUserID)
            frmStatus.txtStatus.SelText = "-" & vbCrLf
            DoEvents
            socket.Close
            Exit For
        End If
    Next socket
End Sub

Private Sub ISON_Timer()
    If notifylist = "" Then
        ISON.Enabled = False
        DoColor frmStatus.txtStatus, "" & color.notify & "*** Notification of users online has been disabled, you have no contacts on your list"
    Else
        If tcp.State = sckConnected Then
            tcp.SendData "ISON " & RTrim(notifylist) & vbCrLf
        End If
    End If
End Sub

Private Sub MDIForm_Load()

    iniFilename = "irc.ini"
    Me.AutoShowChildren = False
    Dim iOption As String
    
    'show splash or not
    iOption = ReadINI("options", "splash")
    If iOption <> "1" Then
        frmSplash.SHOW 1, Me
    End If
    
    'show status window
    Load frmStatus
    frmStatus.SHOW
    
    'show contacts
    'frmContacts.SHOW 0, mdiMain

    
    'show options screen at startup?
    iOption = ReadINI("SHOW", "OPTIONS")
    If iOption Then
        frmOptions.SHOW 0, mdiMain
    End If
    
    'Will you be NOTIFIED of your contacts?
    If iOption <> "" Then
        If iOption Then
            mdiMain.ISON.Enabled = True
        Else
            mdiMain.ISON.Enabled = False
        End If
    End If
    
    
    connected = False
    
    'setup identd sockets
    'load dcc chat sockets
    For i = 1 To maxtcp
        Load ident(i)
        Load CHATx(i) 'used to create chats or send files
        Load CHAT(i) 'used to accept chat or files
    Next i
    
   'load ini options to client
   '[SHOW]
    'quits = 0
    'joinpart = 0
    'modes = 1
    'topics = 1
    'Options = 1
    'kicks = 1
    'CHANFOLDER = 1
    iShow.motd = ReadINI("IRC", "SHOWMOTD")
    iOption = ReadINI("IRC", "SKIPMOTD")
    iShow.joinpart = ReadINI("SHOW", "JOINPART")
    iShow.channelfolder = ReadINI("SHOW", "CHANFOLDER")
    iShow.kicks = ReadINI("SHOW", "KICKS")
    iShow.topics = ReadINI("SHOW", "TOPICS")
    iShow.quits = ReadINI("SHOW", "QUITS")
    iShow.modes = ReadINI("SHOW", "MODES")
    iShow.address = ReadINI("IRC", "SHOWADDRESS")
    iShow.whoisnotify = ReadINI("IRC", "WHOISNOTIFY")
    iShow.notifylist = ReadINI("ISON", "LIST")
    
    'add MyModes "+"
    MyModes = "+"
    
    'background color BULLSHIT
    charf.dwMask = CFM_BACKCOLOR
    charf.cbSize = LenB(charf) 'setup the size of the character format
    
    
    
    'load user set colors
    Dim colorX(0 To 15) As Long
    colorX(0) = vbWhite 'white
    colorX(1) = vbBlack 'black
    colorX(2) = RGB(0, 0, 140) 'dark blue
    colorX(3) = RGB(0, 140, 0) 'dark green
    colorX(4) = vbRed 'red
    colorX(5) = RGB(110, 65, 0) 'brown
    colorX(6) = RGB(140, 0, 140) 'purple
    colorX(7) = RGB(248, 146, 0) 'orange
    colorX(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    colorX(9) = vbGreen 'light green
    colorX(10) = RGB(0, 140, 140) 'dark blue green
    colorX(11) = RGB(0, 255, 255) 'light blue green
    colorX(12) = vbBlue 'light blue
    colorX(13) = vbMagenta 'magenta
    colorX(14) = RGB(140, 140, 140) 'grey
    colorX(15) = RGB(200, 200, 200) 'light grey

    
    Dim srColor As String
    srColor = ReadINI("IRC", "COLORS")
    Dim getColors() As String
    Dim TagColor() As String
    getColors = Split(srColor, " ")
    For i = 0 To UBound(getColors)
        TagColor = Split(getColors(i), ":")
        Select Case i
            Case 0
                color.bgText = TagColor(1)
                frmStatus.txtSend.BackColor = colorX(TagColor(1))
                frmStatus.txtStatus.BackColor = colorX(TagColor(1))
            Case 1
                color.normal = TagColor(1)
                frmStatus.txtSend.ForeColor = colorX(TagColor(1))
            Case 2
                color.ctcp = TagColor(1)
            Case 3
                color.notice = TagColor(1)
            Case 4
                color.action = TagColor(1)
            Case 5
                color.invite = TagColor(1)
            Case 6
                color.join = TagColor(1)
            Case 7
                color.kick = TagColor(1)
            Case 8
                color.mode = TagColor(1)
            Case 9
                color.nick = TagColor(1)
            Case 10
                color.notify = TagColor(1)
            Case 11
                color.part = TagColor(1)
            Case 12
                color.quit = TagColor(1)
            Case 13
                color.topic = TagColor(1)
            Case 14
                color.whois = TagColor(1)
            Case 15
                color.server = TagColor(1)
        End Select
    Next i
    
    'set file index for dcc transfers
    FileIndex = 3
    'set for dcc sends
    FileListenPort = 1559
    
    
    'channel shit
    For i = 1 To ChannelMax
        ChannelModes(i) = ""
    Next i
    
    
    'setup left main treeview
    tvMain.FullRowSelect = True
    tvMain.Nodes.Add , , "mainFriends", "Friends", 2
    tvMain.Nodes.Add , , "mainChats", "Chats", 5
    tvMain.Nodes.Add , , "mainChannels", "Channels", 2
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim i As Integer
    'max sockets to use for ident
    For i = 1 To maxtcp
        Unload ident(i)
        Unload CHATx(i)
        Unload CHAT(i)
    Next i

    Unload frmChannels
    DisableURLDetect
    End
End Sub

Private Sub MDIForm_Resize()
    txtRaw.Move txtRaw.left, txtRaw.top, Me.Width - 350, txtRaw.Height
End Sub

Private Sub mnuArrangeIcons_Click()
    mdiMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    mdiMain.Arrange vbCascade
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGeneralOptions_Click()
    frmOptions.SHOW 0, mdiMain
    
End Sub



Private Sub mnuRaw_Click()

    If mnuRaw.Checked = True Then
        mnuRaw.Checked = False
        'Unload frmRAW
        txtRaw.Text = ""
        picRaw.Visible = False
    Else
        mnuRaw.Checked = True
        'frmRAW.SHOW
        picRaw.Visible = True
    End If
End Sub

Private Sub mnuTile_Click()
    mdiMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuWeb_Click()
    'Load frmWeb
End Sub

Private Sub picHolder_Resize()
    cReBar1.RebarSize
   picHolder.Height = cReBar1.RebarHeight * Screen.TwipsPerPixelY
End Sub

Private Sub script_Error()
    DoColor frmStatus.txtStatus, "" & color.normal & "Error in Script" & vbCrLf & vbCrLf & _
           "Line    :" & script.Error.Line & vbCrLf & _
           "Column  :" & script.Error.Column & vbCrLf & _
           "Source  :" & script.Error.Text & vbCrLf & vbCrLf & _
           script.Error.Description
End Sub

Private Sub picSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrag = True
End Sub

Private Sub picSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    'Adjust the width if the user is holding down the mouse button
    If bDrag = True Then
        tvMain.Height = picSize.top - 150
        picSize.top = Y + picSize.top
        tvMain.Width = picInfo.Width - 150
        tvNotify.top = picSize.top + 150
    End If
End Sub

Private Sub picSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrag = False
End Sub

Private Sub picVBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrag = True
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    'Adjust the width if the user is holding down the mouse button
    If bDrag = True Then
        picInfo.Width = X + picInfo.Width
        picSize.Width = picInfo.Width
        tvMain.Width = picInfo.Width - 150
        tvNotify.Width = picInfo.Width - 150
    End If
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDrag = False
End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
    '.Bevel = sbrRaised 'sbrInset
    'Panel - name of panel clicked
    'StatusBar.Panels.Count - how many panels shown
    'StatusBar.Panels.Item(1).Key - string of panel
    Dim i As Integer
    Dim X As Integer
    Dim re As Integer
    For i = 1 To StatusBar.Panels.Count
        If StatusBar.Panels.Item(i).Key = Panel Then
            StatusBar.Panels.Item(i).Bevel = sbrInset
            'show status
            If LCase(Panel) = "status" Then
                 re = ShowWindow(frmStatus.hwnd, SW_RESTORE)
                 frmStatus.SetFocus
            End If
            'show channel
            If left(Panel, 1) = "#" Then
                For X = 1 To ChannelMax
                    If LCase(ChannelName(X)) = LCase(Panel) Then
                        If LCase(Me.ActiveForm.Caption) = LCase(channel(X).Caption) Then
                            re = ShowWindow(channel(X).hwnd, SW_MINIMIZE)
                        Else
                            re = ShowWindow(channel(X).hwnd, SW_RESTORE)
                            channel(X).SetFocus
                            channel(X).WindowState = 0
                        End If
                    End If
                Next X
            Else
                'show query
                For X = 1 To 100
                    If LCase(QueryName(X)) = LCase(Panel) Then
                        If LCase(Me.ActiveForm.Caption) = LCase(Query(X).Caption) Then
                            re = ShowWindow(Query(X).hwnd, SW_MINIMIZE)
                        Else
                            re = ShowWindow(Query(X).hwnd, SW_RESTORE)
                            Query(X).SetFocus
                            Query(X).WindowState = 0
                        End If
                    End If
                Next X
            End If
        Else
            StatusBar.Panels.Item(i).Bevel = sbrRaised
        End If
    Next i
End Sub



Private Sub tcp_Close()
    Dim i As Integer
    
    connected = False
    DoColor frmStatus.txtStatus, "2*** Disconnected from server!" & vbCrLf & "-" & vbCrLf
    MyModes = ""
    frmStatus.Caption = "Cabral Status: "
    
    For i = 1 To ChannelMax
        Unload channel(i)
        ChannelName(i) = ""
    Next i
    
    'frmFriends.lstFriends.Clear
End Sub

Private Sub tcp_Connect()
    MyModes = ""
    connected = True
    DoColor frmStatus.txtStatus, "2*** Connected to server!" & vbCrLf & "-" & vbCrLf
    tcp.SendData "User " & email & " " & tcp.LocalHostName & " " & tcp.RemoteHost & " :" & RealName & vbCrLf
    tcp.SendData "NICK " & nickname & vbCrLf
    frmStatus.Caption = "Cabral Status: " & nickname
    
    For i = 1 To ChannelMax
        Unload channel(i)
        ChannelName(i) = ""
    Next i
    
    'If iShow.notifylist = 1 Then frmFriends.SHOW
End Sub

Private Sub tcp_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    PingReply = GetTickCount

    
        
    
    Dim strData As String
    Static RestLine As String
    Dim IRCLine() As String
    
    tcp.GetData strData
    
    'Raw information sent to raw window if window is open
    If mnuRaw.Checked = True Then
        'frmRAW.txtRAW = strData
        'frmRAW.txtRaw = tmpString
        txtRaw = txtRaw & strData
    End If
    
    
    If Len(RestLine) > 0 Then
        strData = RestLine & strData
    End If
    
    'IRCLine = Split(strData, vbCrLf)
    IRCLine = Split(strData, Chr(10))
    
    
    If Right(strData, 1) = Chr(10) Or Right(strData, 1) = Chr(13) Then
        RestLine = ""
        For i = 0 To UBound(IRCLine)
            If IRCLine(i) <> "" Then
                IRCLine(i) = Replace(IRCLine(i), Chr(22), "")
                IRCLine(i) = Replace(IRCLine(i), Chr(13), "")
                IRCLine(i) = Replace(IRCLine(i), Chr(10), "")
                CheckWord IRCLine(i)
                'colString.Add IRCLine(i)
            End If
        Next i
    Else
        RestLine = IRCLine(UBound(IRCLine))
        For i = 0 To UBound(IRCLine) - 1
            If IRCLine(i) <> "" Then
                IRCLine(i) = Replace(IRCLine(i), Chr(22), "")
                IRCLine(i) = Replace(IRCLine(i), Chr(13), "")
                IRCLine(i) = Replace(IRCLine(i), Chr(10), "")
                CheckWord IRCLine(i)
                'colString.Add IRCLine(i)
                'DoEvents
            End If
        Next i
    End If

End Sub

Private Sub tcp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DoColor frmStatus.txtStatus, "2*** Error [" & Number & "]: " & Description & vbCrLf & "-" & vbCrLf
    MyModes = ""
    frmStatus.Caption = "Cabral Status: "
    
    Dim i As Integer
    For i = 1 To ChannelMax
        Unload channel(i)
        ChannelName(i) = ""
    Next i
    frmFriends.lstFriends.Clear
End Sub






Private Sub timeParse_Timer()
    'DoColor frmStatus.txtStatus, colString.Count

    Dim word() As String
    Dim strType As String
    Dim parms As String
    Dim intcount As Integer
    
    'split the commands into seperate words
    
    'ReDim Preserve statement is the KEY
    'Do Until InStr(strWord, Chr(32)) = 0
    
    'For intcount = 1 To colString.Count
    '    DoColor frmStatus.txtStatus, colString.Item(intcount)
    'Next intcount
    intcount = 1
    Do Until intcount > colString.Count
        CheckWord colString.Item(intcount)
        colString.Remove intcount
        intcount = intcount + 1
    Loop


End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    'MsgBox Button.Index
    Select Case Button.Index
        Case "1"
            'LOAD server options
            Dim strServer As String
            Dim strPort As String
            strServer = ReadINI("INFO", "SERVER")
            strPort = ReadINI("INFO", "PORT")
            DoColor frmStatus.txtStatus, "" & color.server & "Attempting to connect to " & strServer & " on port " & strPort
            mdiMain.tcp.Close
            mdiMain.tcp.Connect strServer, Val(strPort)
            
            nickname = ReadINI("INFO", "NICKNAME")
            email = ReadINI("INFO", "USERNAME")
            RealName = ReadINI("INFO", "REALNAME")

        Case "2"

        Case "3"
            'options
            frmOptions.SHOW 0, mdiMain
        Case "4"
            'fave channels
            frmChannelFolder.SHOW 0, mdiMain
        Case "5"
            '
        Case "6"
            'scripting
            frmScript.SHOW
        Case "7"
            'list channels
            frmChanList.SHOW 0, mdiMain
            'notify list
            'frmNotify.SHOW
        Case "8"
            'url catcher
            frmURL.SHOW
        Case "9"
        Case "10"
            'send file
            'frmSendFile.SHOW 0, mdiMain
            Dim NewSendFileWin As frmSendFile
            Set NewSendFileWin = New frmSendFile
            FileListenPort = FileListenPort + 1
            If FileListenPort > 9000 Then FileListenPort = 1560
            Load NewSendFileWin.tcpSend(FileListenPort)
            NewSendFileWin.Tag = Str(FileListenPort)
            
            NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort = NewSendFileWin.Tag
            'MsgBox NewSendFileWin.tcpSend(NewSendFileWin.Tag).LocalPort
            NewSendFileWin.tcpSend(NewSendFileWin.Tag).Listen
            NewSendFileWin.SHOW 0, Me
            
            
            
        Case "11"
            'dcc chat
            frmDCCCHAT.SHOW 0, mdiMain
        Case "12"
        Case "13"
            'Tile
            mdiMain.Arrange vbTileHorizontal
        Case "14"
            'cascade
            mdiMain.Arrange vbCascade
        Case "15"
        Case "16"
            'help
    End Select
End Sub

Private Sub tvMain_NodeClick(ByVal Node As MSComctlLib.Node)
    'Node.Key
    On Error Resume Next
If Node.Key <> LCase(mainchannels) Or Node.Key <> LCase(mainchats) Or Node.Key <> LCase(mainfriends) Then
    For i = 1 To 100
        If LCase(Node.Key) = LCase(ChannelName(i)) Or LCase(Node.Key) = LCase(QueryName(i)) Then
            channel(i).SetFocus
            Exit Sub
        End If
    Next i
End If
End Sub

Private Sub txtRaw_Change()
    txtRaw.SelStart = Len(txtRaw)
End Sub

Private Sub txtSearch_GotFocus()
    If txtSearch = "<search the web>" Then
        txtSearch = ""
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    Dim webPage As Integer
    If KeyAscii = 13 Then
        webPage = ShellExecute(Me.hwnd, "Open", txtSearch, "", App.Path, 1)
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtSearch_LostFocus()
    If Trim(txtSearch) = "" Then
        txtSearch = "<search the web>"
    End If
End Sub
