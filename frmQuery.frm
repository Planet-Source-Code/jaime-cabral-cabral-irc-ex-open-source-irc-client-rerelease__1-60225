VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmQuery 
   Caption         =   "Private Message Window"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   5520
   Begin RichTextLib.RichTextBox txtQuery 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
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
      TextRTF         =   $"frmQuery.frx":1042
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
      TabIndex        =   0
      Top             =   2280
      Width           =   4575
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call Form_Resize
    
    Me.Height = 5000
    Me.Width = 6000
    EnableURLDetect txtQuery.hwnd, Me.hwnd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    txtQuery.Move txtQuery.left, txtQuery.top, Me.Width - 150, Me.Height - 700
    txtSend.Move txtQuery.left, txtQuery.Height - 13, Me.Width - 150, 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim word() As String
    word = Split(Me.Caption, " ")
    
    For i = 1 To 100
        If LCase(word(0)) = LCase(QueryName(i)) Then
            QueryName(i) = ""
            Call RemoveTaskbar(word(0))
            RemoveNode word(0)
        End If
    Next i
End Sub

Private Sub txtQuery_Change()
    txtQuery.SelStart = Len(txtQuery.Text)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    Dim word() As String
    word = Split(Me.Caption, Chr(32))
    If KeyAscii = 13 Then
        If left(txtSend.Text, 1) = "/" Then
            ACTION_CHANNEL = word(0)
            Call xINPUT(Mid(txtSend, 2), Me.txtQuery)
        Else
            mdiMain.tcp.SendData "PRIVMSG " & word(0) & " :" & txtSend & vbCrLf
            'DoColor txtQuery, "" & color.normal & "<" & color.action & nickname & "" & color.normal & "> " & txtSend
            DoColor txtQuery, "" & color.normal & "<" & color.action & "" & nickname & "" & color.normal & "> " & txtSend
        End If
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
