VERSION 5.00
Begin VB.Form frmAddNetwork 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Network"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmAddNetwork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtNetwork 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Network Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1110
   End
End
Attribute VB_Name = "frmAddNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    'addnetwork
    'add network to "allgroups"
    frmOptions.cmbNetwork.AddItem txtNetwork
    frmOptions.cmbNetwork.Text = txtNetwork
    
    Dim word() As String
    Dim varX As String
    Dim i As Integer
    iniFilename = "servers.ini"
    frmOptions.lvwServers.ListItems.Clear
    For i = 0 To 500
        varX = i
        OutPut = ReadINI(frmOptions.cmbNetwork.Text, varX)
        'If OutPut = "" Then Exit For
        If OutPut <> "" Then
            word = Split(OutPut, "|")
            Set mItem = frmOptions.lvwServers.ListItems.Add(, , word(1))
            mItem.SubItems(1) = word(0)
            mItem.SubItems(2) = word(2)
        End If
    Next i

    
    
    Unload Me
End Sub
