VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChannel 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4635
   Icon            =   "channel.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   4635
   Begin VB.PictureBox picMSG 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3240
      Picture         =   "channel.frx":1042
      ScaleHeight     =   120
      ScaleMode       =   0  'User
      ScaleWidth      =   120
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picVOICE 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2880
      Picture         =   "channel.frx":1386
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picUNVOICE 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2520
      Picture         =   "channel.frx":16CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picUNOP 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2160
      Picture         =   "channel.frx":1A0E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picTOPIC 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1800
      Picture         =   "channel.frx":1D52
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picOP 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1440
      Picture         =   "channel.frx":2096
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picKICK 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1080
      Picture         =   "channel.frx":23DA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picCHANNEL 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   720
      Picture         =   "channel.frx":271E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picBAN 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   360
      Picture         =   "channel.frx":2A62
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picJOIN 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   0
      Picture         =   "channel.frx":2DA6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin RichTextLib.RichTextBox txtText 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"channel.frx":30EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      IntegralHeight  =   0   'False
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox txtTopic 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"channel.frx":3165
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuChannelNicks 
      Caption         =   "Cabral"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois ?"
      End
   End
End
Attribute VB_Name = "frmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Form_Activate()
    Dim word() As String
    word = Split(Me.Caption)
    UNhighlight_node word(0)
End Sub

Private Sub Form_Load()
    Call Form_Resize
    
    Me.Height = 5000
    Me.Width = 6000

    Me.Hide
    EnableURLDetect txtText.hwnd, Me.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'Dim i As Integer
    'For i = 1 To 10
    '    If LCase(Me.Caption) = ChannelName(i) Then
            'txtText.Move txtText.left, txtTopic.Height - 13 , Me.Width - 1500, Me.Height - txtSend.Height - 700
            txtText.Move txtText.left, 0, Me.Width - 1600, Me.Height - txtSend.Height - 420
            'lstNames.Move txtText.Width - 13, txtText.top, 1400, txtText.Height
            lstNames.Move txtText.Width + 10, txtText.top, 1465, txtText.Height
            txtSend.Move txtText.left, txtText.Height + 10, Me.Width - 125, 350
            txtTopic.Move txtText.left, txtTopic.top, txtText.Width + lstNames.Width - 25, 300
            txtTopic.Visible = False
    '    End If
    'Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    

    On Error Resume Next
    Dim i As Integer
    Dim word() As String
    word = Split(Me.Caption, " ")
    mdiMain.tcp.SendData "PART " & word(0) & vbCrLf
    For i = 1 To ChannelMax
        If LCase(ChannelName(i)) = LCase(word(0)) Then
            ChannelName(i) = ""
            ChannelModes(i) = ""
            ChannelTopic(i) = ""
            'remove from status bar
            Call RemoveTaskbar(word(0))
            RemoveNode word(0)
        End If
    Next i
End Sub



Private Sub lstNames_DblClick()
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





    Dim i As Integer
    Dim xlfound As Boolean
    xlfound = False
    Dim ChatWith As String
    ChatWith = lstNames.List(lstNames.ListIndex)
    ChatWith = Replace(ChatWith, "@", "")
    ChatWith = Replace(ChatWith, "+", "")
    ChatWith = Replace(ChatWith, "%", "")

    For i = 1 To 100
        If LCase(QueryName(i)) = LCase(ChatWith) Then
            xlfound = True
            Exit For
        End If
    Next i
    
If xlfound = False Then
    For i = 1 To 100
        If QueryName(i) = "" Then
            Load Query(i)
            'set user colors
            Query(i).txtSend.BackColor = colorX(color.bgText)
            Query(i).txtSend.ForeColor = colorX(color.normal)
            Query(i).txtQuery.BackColor = colorX(color.bgText)
            '
            Query(i).Caption = ChatWith
            QueryName(i) = ChatWith
            Call AddTaskbar(ChatWith, 1)
            mdiMain.tvMain.Nodes.Add "mainChats", tvwChild, LCase(ChatWith), ChatWith, 3
            Exit For
        End If
    Next i
End If
End Sub

Private Sub lstNames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRetVal As Variant
    If Button = vbRightButton Then
    lRetVal = SendMessage(Me.hwnd, WM_RBUTTONDOWN, 0, 0)
        Call PopupMenu(Me.mnuChannelNicks)
    End If
End Sub

Private Sub mnuWhois_Click()
    Dim strWhois As String
    strWhois = lstNames.List(lstNames.ListIndex)
    strWhois = Replace(strWhois, "@", "")
    strWhois = Replace(strWhois, "+", "")
    mdiMain.tcp.SendData "WHOIS " & strWhois & vbCrLf
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    Dim word() As String
    word = Split(Me.Caption, Chr(32))
    If KeyAscii = 13 Then
        txtSend = LTrim(txtSend)
        If left(txtSend.Text, 1) = "/" Then
            ACTION_CHANNEL = LCase(word(0))
            Call xINPUT(Mid(txtSend, 2), Me.txtText)
        Else
            mdiMain.tcp.SendData "PRIVMSG " & word(0) & " :" & txtSend & vbCrLf
            DoColor txtText, "" & color.normal & "<" & color.action & "" & nickname & "" & color.normal & "> " & txtSend
        End If
        'txtText.SelText = txtSend & vbCrLf
        txtSend = ""
        KeyAscii = 0
    End If
    If KeyAscii = 11 Then
        'places the color code box
        '
        'txtSend.SelStart = where you are gonna insert.
        Dim starttext As Integer
        starttext = txtSend.SelStart
        txtSend = Mid(txtSend, 1, txtSend.SelStart) & "" & Mid(txtSend, txtSend.SelStart + 1)
        txtSend.SelStart = starttext + 1
    End If
End Sub


Private Sub txtText_Change()
        'txtText.SelStart = Len(txtText)
        If Len(txtText) > 30000 Then
            LockWindowUpdate txtText.hwnd
            txtText.SelStart = 0
            txtText.SelLength = Len(txtText) - 20000
            txtText.SelText = ""
            LockWindowUpdate 0
        End If
        txtText.SelStart = Len(txtText)
        
        Dim word() As String
        word = Split(Me.Caption)
        If mdiMain.ActiveForm.Caption <> Me.Caption Then
            highlight_node word(0)
        End If
End Sub

Private Sub txtText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.SetText txtText.SelText
    Me.txtSend.SetFocus
End Sub

Private Sub txtText_SelChange()
    'txtText.SelStart = Len(txtText)
End Sub

Private Sub txtTopic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mdiMain.tcp.SendData "TOPIC " & Me.Caption & " :" & txtTopic.Text & vbCrLf
    End If
End Sub
