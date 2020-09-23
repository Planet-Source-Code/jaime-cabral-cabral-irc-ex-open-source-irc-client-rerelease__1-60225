VERSION 5.00
Begin VB.Form frmAddNickName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cabral Add NickName"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "frmAddNickName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbsave 
      Caption         =   "OK"
      Height          =   435
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmbRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmbAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtNickName 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.ListBox lstNickname 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Nickname to use:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmAddNickName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAdd_Click()
    If txtNickName.Text <> "" Then
        lstNickname.AddItem txtNickName.Text
    End If
End Sub

Private Sub cmbRemove_Click()
    On Error Resume Next
    lstNickname.RemoveItem lstNickname.ListIndex
End Sub

Private Sub cmbsave_Click()
    CNumber = FreeFile
    Open App.Path & "\names.ini" For Output As #CNumber
    For X = 0 To lstNickname.ListCount - 1
        Print #CNumber, lstNickname.List(X)
    Next X
    Close #CNumber
    Unload Me
End Sub

Private Sub Form_Load()
    CNumber = FreeFile
    'read nickname file
    Open App.Path & "\names.ini" For Input As #CNumber
    Do While Not (EOF(CNumber))
        Line Input #CNumber, i
        lstNickname.AddItem i
    Loop
    Close #CNumber
End Sub

Private Sub lstNickname_Click()
    'txtNickName = lstNickname.ListIndex
    txtNickName = lstNickname.List(lstNickname.ListIndex)
End Sub
