VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmMOTD 
   Caption         =   "Cabral IRC:  Message of the Day"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "ss.frx":0000
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
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"ss.frx":038A
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    txtMOTD.Move txtMOTD.Left, txtMOTD.Top, Me.Width - 150, Me.Height - 500
End Sub

Private Sub txtMOTD_LinkOver(ByVal iType As vbalEdit.ERECLinkEventTypeCOnstants, ByVal lMin As Long, ByVal lMax As Long)
   If (iType = ercLButtonUp) Then
      MsgBox "Use ShellEx to run this shortcut: " & txtMOTD.TextInRange(lMin, lMax), vbInformation
   Else
      'MsgBox "LinkOver: " & txtMOTD.TextInRange(lMin, lMax)
   End If
End Sub

