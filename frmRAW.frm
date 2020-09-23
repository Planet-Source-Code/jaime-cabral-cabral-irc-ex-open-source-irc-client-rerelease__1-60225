VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmRAW 
   Caption         =   "Cabral IRC: Raw Incoming Server Text"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "frmRAW.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5265
   Begin RichTextLib.RichTextBox txtRAW 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      _Version        =   327681
      TextRTF         =   $"frmRAW.frx":08CA
   End
End
Attribute VB_Name = "frmRAW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    txtRAW.Move txtRAW.Left, txtRAW.Top, Me.Width - 150, Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiMain.mnuRaw.Checked = False
End Sub

Private Sub txtRaw_Change()
    txtRAW.SelStart = Len(txtRAW)
End Sub
