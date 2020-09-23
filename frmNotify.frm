VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNotify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cabral Notify"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin TabDlg.SSTab tabNotify 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Notify"
      TabPicture(0)   =   "frmNotify.frx":1AFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSTab1_DblClick()

End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer
    If Trim(txtNickname) <> "" Then
        i = CheckListbox(lstNotify, Trim(txtNickname))
        If i = 0 Then
            lstNotify.AddItem Trim(txtNickname)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    For i = 0 To lstNotify.ListCount - 1
        If lstNotify.List(lstNotify.ListIndex) = lstNotify.List(i) Then
            lstNotify.RemoveItem i
            Exit For
        End If
    Next i
End Sub

Private Sub cmdOK_Click()
    'save nicknames to file
    'and add to array
    Dim i As Integer
    For i = 1 To 100
        notify(i) = ""
    Next i
    NOTIFYLIST = ""
    Open App.Path & "\notify.txt" For Output As #1
    For i = 0 To lstNotify.ListCount - 1
        Print #1, lstNotify.List(i)
        notify(i + 1) = lstNotify.List(i)
        NOTIFYLIST = NOTIFYLIST & notify(i + 1) & " "
    Next i
    Close #1
    If chkEnable = 1 Then
        mdiMain.ISON.Enabled = True
    Else
        mdiMain.ISON.Enabled = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    'load USERS
    Dim USERS As String
    If FileExists(App.Path & "\notify.txt") Then
        CNumber = FreeFile
        'read nickname file
        Open App.Path & "\notify.txt" For Input As #CNumber
        Do While Not (EOF(CNumber))
            Line Input #CNumber, i
            If Trim(i) <> "" Then
                lstNotify.AddItem i
            End If
        Loop
        Close #CNumber
    End If
End Sub
