VERSION 5.00
Begin VB.Form frmChannelFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cIRC Channel Folder"
   ClientHeight    =   4065
   ClientLeft      =   6225
   ClientTop       =   3405
   ClientWidth     =   4680
   Icon            =   "frmChannelFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox chkShowChanFolder 
      Caption         =   "Pop up folder on connect"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3780
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.ListBox lstChannels 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtChannel 
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
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter name of channel to join:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmChannelFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    If left(txtChannel, 1) <> "#" Then
        txtChannel = "#" & txtChannel
    End If
    txtChannel = Trim(txtChannel)
    lstChannels.AddItem txtChannel
End Sub

Private Sub cmdJoin_Click()
    MsgBox lstChannels.List(lstChannels.ListIndex)
    mdiMain.tcp.SendData "JOIN " & lstChannels.List(lstChannels.ListIndex) & vbCrLf
End Sub

Private Sub cmdOK_Click()
    CNumber = FreeFile
    Open App.Path & "\chanfolder.ini" For Output As #CNumber
    For X = 0 To lstChannels.ListCount - 1
        Print #CNumber, lstChannels.List(X)
    Next X
    Close #CNumber
    
    i = WriteINI("SHOW", "CHANFOLDER", chkShowChanFolder.Value)
    
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    lstChannels.RemoveItem lstChannels.ListIndex
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim fe As Boolean
    fe = FileExists(App.Path & "\chanfolder.ini")
    If fe Then
        CNumber = FreeFile
        'read servers file
        Open App.Path & "\chanfolder.ini" For Input As #CNumber
        Do While Not (EOF(CNumber))
            Line Input #CNumber, i
            lstChannels.AddItem i
        Loop
        Close #CNumber
    End If
    
    chkShowChanFolder.Value = ReadINI("SHOW", "CHANFOLDER")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    chkShowChanFolder.Value = ReadINI("SHOW", "CHANFOLDER")
    iShow.channelfolder = chkShowChanFolder.Value
End Sub

Private Sub lstChannels_DblClick()
    On Error Resume Next
    If left(lstChannels.List(lstChannels.ListIndex), 1) = "#" Then
        mdiMain.tcp.SendData "JOIN " & lstChannels.List(lstChannels.ListIndex) & vbCrLf
    Else
        mdiMain.tcp.SendData "JOIN #" & lstChannels.List(lstChannels.ListIndex) & vbCrLf
    End If
End Sub
