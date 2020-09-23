VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Cabral DCC Chat: "
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   Icon            =   "frmDCC.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5445
   Begin RichTextLib.RichTextBox txtDCC 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDCC.frx":1042
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
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   5415
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call DoColor(Me.txtDCC, "4* DCC chat with " & Me.Caption)
    EnableURLDetect txtDCC.hwnd, Me.hwnd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    txtDCC.Move txtDCC.left, txtDCC.top, Me.Width - 150, Me.Height - 700
    txtSend.Move txtDCC.left, txtDCC.Height - 10, Me.Width - 150, 300
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    For i = 1 To UBound(ChatWindowName)
        If LCase(Me.Caption) = LCase(ChatWindowName(i)) Then
            MsgBox ChatWindowName(i) & i
            Unload mdiMain.CHAT(i)
            ChatWindowName(i) = ""
        End If
    Next i
    
    For i = 1 To 25
        If LCase(Me.Caption) = LCase(ChatWindowNamex(i)) Then
            MsgBox ChatWindowNamex(i) & i
            Unload mdiMain.CHATx(i)
            ChatWindowNamex(i) = ""
        End If
    Next i
End Sub

Private Sub txtDCC_Change()
    txtDCC.SelStart = Len(txtDCC)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        For i = 1 To maxtcp
            'MsgBox ChatWindowName & " " & Me.Caption
            If LCase(ChatWindowName(i)) = LCase(Me.Caption) Then
                mdiMain.CHAT(i).SendData txtSend & vbCrLf
                'Me.txtDCC.SelText = "<" & nickname & "> " & txtSend & vbCrLf
                DoColor Me.txtDCC, "0<" & nickname & "> " & txtSend
                txtSend = ""
            End If
        Next i
        For i = 1 To maxtcp
            If LCase(ChatWindowNamex(i)) = LCase(Me.Caption) Then
                'MsgBox "yeah"
                mdiMain.CHATx(i).SendData txtSend & vbCrLf
                'Me.txtDCC.SelText = "<" & nickname & "> " & txtSend & vbCrLf
                DoColor Me.txtDCC, "0<" & nickname & "> " & txtSend
                txtSend = ""
            End If
        Next i
    End If
End Sub
