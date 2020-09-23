VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCCFILE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cabral DCC GET"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmDCCFILE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock FILE 
      Index           =   0
      Left            =   1080
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   2805
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   4948
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox picComplete 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   13
      Top             =   240
      Width           =   135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdGET 
      Caption         =   "GET"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Label lblAddress 
         Caption         =   "Label4"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblFileSize 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblFile 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblNickName 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NickName:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   285
      End
   End
   Begin VB.Label lblRCV 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label lblFilename 
      Caption         =   "Save Path"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
End
Attribute VB_Name = "frmDCCFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload FILE(Me.Tag)
    Close #Me.Tag
    Unload Me
End Sub

Private Sub cmdGET_Click()
    Dim cc() As String
    cc = Split(lblAddress, ":")
    'MsgBox cc(0) & ":" & cc(1)
    FILE(Me.Tag).Close
    FILE(Me.Tag).Connect cc(0), cc(1)
    cmdGET.Enabled = False
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub FILE_Close(Index As Integer)
    Call DoColor(frmStatus.txtStatus, "4* File Connection Closed! [" & Me.Tag & "]")
    Close #Me.Tag
End Sub

Private Sub FILE_Connect(Index As Integer)
    Call DoColor(frmStatus.txtStatus, "4* Ready To Recieve File!")
    'Open frmDCCFILE.lblFilename For Binary As ffile
    'ffile = FileIndex + 1
    Dim fname As String
    fname = App.Path & "\downloads\"
    'MkDir fname
    Open App.Path & "\" & Me.lblFile For Binary As Me.Tag
    'Open App.Path & "\" & frmDCCFILE.lblFile For Output As Index


End Sub

Private Sub FILE_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Me.ProgressBar.Value = Me.ProgressBar.Value + bytesTotal
    Me.lblRCV = "Recieved: " & Me.ProgressBar.Value
    
    Dim ReadBuffer() As Byte
    Dim Retval As Long
    FILE(Me.Tag).GetData ReadBuffer, vbByte
    Put #Me.Tag, , ReadBuffer

    'send back acknowledgement
    Hexdata = Hex(LOF(Me.Tag))
    Hexdata = String$(8 - Len(Hexdata), "0") & Hexdata
    ReDim SendBackData(3) As Byte
    For i = 1 To Len(Hexdata) Step 2
        SendBackData((i - 1) / 2) = Val("&H" & Mid(Hexdata, i, 2))
    Next
    FILE(Me.Tag).SendData SendBackData

    If Me.ProgressBar.Value = Me.ProgressBar.Max Then
        Me.lblRCV = "File Transfer Complete"
        Me.picComplete.BackColor = vbRed
    'Else
    '    Me.picComplete.BackColor = vbWhite
    End If
End Sub

Private Sub FILE_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Close #Me.Tag
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Close #Me.Tag
End Sub
