VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendFile 
   Caption         =   "DCC SEND"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmSendFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "?"
         Height          =   300
         Left            =   3600
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Elapsed &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblTimeElapsed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   19
         Top             =   2400
         Width           =   270
      End
      Begin VB.Label lblSR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sent/Rcvd &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label lblSentRcvd 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   2880
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes/Second &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblBps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   3120
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Left &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label lblTimeLeft 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   2640
         Width           =   270
      End
      Begin VB.Label lblDir 
         Caption         =   "x:\"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dir:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblFileSize 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Send to:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3625
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock tcpSend 
      Index           =   0
      Left            =   480
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStat 
      Alignment       =   1  'Right Justify
      Caption         =   "..."
      Height          =   210
      Left            =   480
      TabIndex        =   21
      Top             =   2400
      Width           =   1920
   End
   Begin VB.Shape picComplete 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFullPath As String
Dim i, fLength, ret                               '// Declare Variables
Dim Buffer As String                              '// Declare Buffer
Dim bSize As Long
Dim ByteSent As Long

Function SendFile() As Boolean
    'i = FreeFile                                      '// Set I As FreeFile
    i = (Val(Me.Tag) - 1530)
    Static pvalue As Integer
    

    
    If lblStat.Caption <> "Sending file" Then
        lblStat.Caption = "Sending file"
    End If
    
    If pvalue <> Int(Progress.Value * 100 / Progress.max) Then
        pvalue = Int(Progress.Value * 100 / Progress.max)
        Me.Caption = "DCC File Transfer: " & pvalue & "%"
    End If
        
    If ByteSent < Progress.max Then
        Progress.Value = ByteSent
    Else
        Progress.Value = Progress.max
    End If
    'Open strFullPath For Binary Access Read As #i
    bSize = 1024
    fLength = LOF(i)
    
    If ByteSent >= fLength Then
        tcpSend(Me.Tag).SendData ""
        Exit Function
    End If
    If ByteSent + bSize > fLength Then
        bSize = fLength - ByteSent
    End If
        
    Buffer = Space$(bSize)
    Get i, , Buffer
    tcpSend(Me.Tag).SendData Buffer
    ByteSent = ByteSent + bSize

    
    If ByteSent >= Val(lblFileSize) Then
        Close #i
        tcpSend(Me.Tag).Close
        Me.Caption = "File Sent"
        picComplete.BackColor = vbRed
    End If
End Function



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    Dim strFile As SelectedFile
    strFile = ShowOpen(Me.hwnd, True)
    If strFile.bCanceled Then Exit Sub
    txtFileName = strFile.sFiles(1)
    lblFileSize = FileLen(txtFileName)
    lblDir = strFile.sLastDirectory
    strFullPath = strFile.sLastDirectory & strFile.sFiles(1)
End Sub

Private Sub cmdSend_Click()
    On Error Resume Next
    Dim LIP As String
    Dim TempFileName As String
    
    'no file selected
    If Trim(txtNickname.Text) = "" Then
        lblStat.Caption = "Please type in a user"
        'cmdSend.Enabled = False
        Exit Sub
    End If
    If Trim(txtFileName.Text) = "" Then
        lblStat.Caption = "Please select a file"
        'cmdSend.Enabled = False
        Exit Sub
    End If
    
    cmdSend.Enabled = False
    lblStat.Caption = "Attempting to connect"
    
    
    TempFileName = Replace(txtFileName, " ", "_")
    LIP = IrcGetLongIP(mdiMain.tcp.LocalIP)
    mdiMain.tcp.SendData "NOTICE " & txtNickname.Text & " :DCC SEND " & txtFileName.Text & "(" & tcpSend(Me.Tag).LocalIP & ")" & vbCrLf
    mdiMain.tcp.SendData "PRIVMSG " & txtNickname & " :DCC SEND " & TempFileName & " " & LIP & " " & tcpSend(Me.Tag).LocalPort & " " & " " & lblFileSize & "" & vbCrLf
    
    Progress.Min = 0
    Progress.max = Val(lblFileSize)
    
End Sub

Private Sub tcpSend_Close(Index As Integer)
    lblStat.Caption = "Connection closed"
    Me.picComplete.BackColor = vbRed
End Sub

Private Sub tcpSend_Connect(Index As Integer)
    MsgBox "COnnected"
End Sub

Private Sub tcpSend_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If tcpSend(Me.Tag).State <> sckClosed Then tcpSend(Me.Tag).Close
    tcpSend(Me.Tag).Close
    tcpSend(Me.Tag).Accept requestID
    Me.Caption = "Accepted connection"
    Me.Caption = "Connection Established"     '// Change Caption
    
    Me.Caption = "Sending..."
    i = (Val(Me.Tag) - 1530)
    Open strFullPath For Binary Access Read As i
    bSize = 1024
    fLength = LOF(i)
                                     
    If fLength - Loc(i) <= bSize Then     '// If The Buffer Is Larger Than
        bSize = fLength - Loc(i)          '// The Rest Of the File. Make h
    End If                                '// New Buffer Size The Rest Of The
                                          '// File
    If bSize = 0 Then Exit Sub             '// If Buffer Size Is 0 Send Done

    ByteSent = ByteSent + bSize           '// Adds The Buffer To Bytes Sent
    Buffer = Space$(bSize)                '// Get The Buffer From The BlockSize
    Get i, , Buffer                         '// Take Block From File
    tcpSend(Me.Tag).SendData Buffer                   '// Send Block
                                   '// Close File
    'Close #i
    
End Sub

Private Sub tcpSend_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    tcpSend(Me.Tag).GetData strData
    'MsgBox "ARRIVAL: " & strData
    
    SendFile
End Sub

Private Sub tcpSend_SendComplete(Index As Integer)
    'MsgBox "SEND COMPLETE"
    
End Sub

Private Sub Form_Load()
    'This is for dcc send
                                   '// Sets Winsock To Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


