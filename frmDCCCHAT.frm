VERSION 5.00
Begin VB.Form frmDCCCHAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DCC Chat"
   ClientHeight    =   1290
   ClientLeft      =   6420
   ClientTop       =   5520
   ClientWidth     =   4680
   Icon            =   "frmDCCCHAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtNAME 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type in the username:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDCCCHAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call xINPUT("CHAT " & txtNAME, frmStatus.txtStatus)
    Unload Me
End Sub
