VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Cabral Status:"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4680
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
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmStatus.frx":1042
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
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()

    EnableURLDetect txtStatus.hwnd, Me.hwnd
    'intro text
    txtStatus.SelBold = True
    txtStatus.SelColor = RGB(140, 0, 0)
    txtStatus.SelText = "Cabral IRC Client" & vbCrLf
    txtStatus.SelColor = vbBlack
    txtStatus.SelText = "Planet Cabral Production" & vbCrLf
    txtStatus.SelBold = False
    txtStatus.SelText = "Please rebort bugs to cabral_jaime@hotmail.com or visit the webpage at www.planetcabral.com -  Have comments, additions, or suggestions?  E-mail me." & vbCrLf
    
    Call AddTaskbar("Status", 5)
    
   

   
   Call Form_Resize
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call RemoveTaskbar("Status")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtStatus.Move txtStatus.left, txtStatus.top, frmStatus.Width - 125, frmStatus.Height - 690
    txtSend.Move txtStatus.left, txtStatus.Height - 13, frmStatus.Width - 125, 300
End Sub

   


Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

    Static LastSent(1 To 20) As String
    Static X As Integer
    'set text = when you press up
    '[chr(38)] it'll show last sent message
    If KeyCode = 38 Then
        X = X + 1
        If X > 10 Then X = 1
        txtSend.Text = LastSent(X)
    End If
    If KeyCode = 40 Then
        txtSend = ""
    End If
        
    If KeyCode = 13 Then
        X = 0
        
        LastSent(10) = LastSent(9)
        LastSent(9) = LastSent(8)
        LastSent(8) = LastSent(7)
        LastSent(7) = LastSent(6)
        LastSent(6) = LastSent(5)
        LastSent(5) = LastSent(4)
        LastSent(4) = LastSent(3)
        LastSent(3) = LastSent(2)
        LastSent(2) = LastSent(1)
        LastSent(1) = txtSend

        For i = 1 To 10
            'DoColor txtStatus, "04" & i & ": " & LastSent(i)
        Next i
        
        If left(txtSend.Text, 1) = "/" Then
            ACTION_CHANNEL = ""
            Call xINPUT(Mid(txtSend, 2), Me.txtStatus)
        End If
        txtSend = ""
    End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    'On Error GoTo xError
    On Error Resume Next

    If KeyAscii = 13 Then
        'mdiMain.tcp.SendData txtSend & vbCrLf
        'txtStatus.SelText = "> " & txtSend & vbCrLf
        txtSend = LTrim(txtSend)
        If left(txtSend.Text, 1) = "/" Then
            Call xINPUT(Mid(txtSend, 2), frmStatus.txtStatus)
        End If


        txtSend = ""
        KeyAscii = 0
    End If

xError:
    If Err.Description <> "" Then
        'MsgBox Err.Description
    End If
End Sub


Private Sub txtStatus_Change()
        If Len(txtStatus) > 20000 Then
            LockWindowUpdate txtStatus.hwnd
            txtStatus.SelStart = 0
            txtStatus.SelLength = Len(txtStatus) - 10000
            txtStatus.SelText = ""
            LockWindowUpdate 0
        End If
        txtStatus.SelStart = Len(txtStatus)
End Sub

Private Sub txtStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.SetText txtStatus.SelText
    txtSend.SetFocus
End Sub
