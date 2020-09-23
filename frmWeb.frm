VERSION 5.00
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "MSHTML.tlb"
Begin VB.Form frmWeb 
   Caption         =   "Cabral IRC Client: open source project"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmWeb.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSHTMLCtl.Scriptlet webOne 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Scrollbar       =   -1  'True
      URL             =   "http://darkimages.cjb.net/"
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    webOne.Move webOne.Left, webOne.Top, Me.Width - 150, Me.Height - 400
End Sub

