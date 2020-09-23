VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMOTD 
   Caption         =   "Cabral IRC:  Message of the Day"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "frmMOTD.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   4620
   Begin RichTextLib.RichTextBox txtMOTD 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMOTD.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    EnableURLDetect txtMOTD.hwnd, Me.hwnd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtMOTD.Move txtMOTD.left, txtMOTD.top, Me.Width - 150, Me.Height - 500
End Sub


Private Sub txtMOTD_Change()
    txtMOTD.SelStart = Len(txtMOTD)
End Sub
