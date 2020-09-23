VERSION 5.00
Begin VB.Form frmEditServer 
   Caption         =   "Edit Server"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   Icon            =   "frmEditServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port Range:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frmEditServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtDescription = frmOptions.lvwServers.SelectedItem.Text
    txtName = frmOptions.lvwServers.SelectedItem.SubItems(1)
    txtPort = frmOptions.lvwServers.SelectedItem.SubItems(2)
End Sub
