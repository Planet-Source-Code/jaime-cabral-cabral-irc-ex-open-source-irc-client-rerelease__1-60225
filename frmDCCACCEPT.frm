VERSION 5.00
Begin VB.Form frmDCCACCEPT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabral DCC Chat"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "frmDCCACCEPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3540
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdChat 
      Caption         =   "Chat!"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.Label lblPort 
         Alignment       =   2  'Center
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblIP 
         Alignment       =   2  'Center
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblNickName 
         Alignment       =   2  'Center
         Caption         =   "NickName"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmDCCACCEPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChat_Click()
    On Error Resume Next
    For i = 1 To maxtcp
        'If mdiMain.CHAT(i).State = sckError Or mdiMain.CHAT(i).State = sckClosed Then
        If ChatWindowName(i) = "" Then
            Unload mdiMain.CHAT(i)
            Load mdiMain.CHAT(i)
            Load ChatWindow(i)
            ChatWindow(i).SHOW
            mdiMain.CHAT(i).Close
            ChatWindow(i).Caption = lblNickName
            ChatWindowName(i) = lblNickName
            mdiMain.CHAT(i).Connect lblIP, lblPort
            Exit For
        End If
    Next i
    Unload Me
End Sub

