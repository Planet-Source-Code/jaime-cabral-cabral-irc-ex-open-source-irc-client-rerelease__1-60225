VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChannels 
   AutoRedraw      =   -1  'True
   Caption         =   "Cabral Chanel List"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "frmChannels.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   6315
   Begin MSComctlLib.ImageList imgChan 
      Left            =   5640
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannels.frx":1042
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwChan 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    ' Clear the ColumnHeaders collection.
    'lvwChan.ColumnHeaders.Clear
    ' Add four ColumnHeaders.
    lvwChan.ColumnHeaders.Add , , "Channel", 3000
    lvwChan.ColumnHeaders.Add , , "Users"
    lvwChan.ColumnHeaders.Add , , "Topic", 3500
    
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvwChan.Move lvwChan.left, lvwChan.top, Me.Width - 100, Me.Height - 400
End Sub


Private Sub lvwChan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwChan.AllowColumnReorder = True
    If ColumnHeader = "Channel" Then
        lvwChan.SortKey = 0
    End If
    If ColumnHeader = "Users" Then
        lvwChan.SortKey = 1
    End If
    If ColumnHeader = "Topic" Then
        lvwChan.SortKey = 2
    End If
End Sub

Private Sub lvwChan_DblClick()
    On Error Resume Next

    If Not (lvwChan.SelectedItem Is Nothing) Then
        'lvwChan_ItemDblClick lvwChan.SelectedItem
        mdiMain.tcp.SendData "JOIN " & lvwChan.SelectedItem & vbCrLf
    End If
End Sub



Private Sub lvwChan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lvwChan.SelectedItem = lvwChan.HitTest(X, Y)
End Sub
Private Sub lvwChan_ItemDblClick(Item As ListItem)
    On Error Resume Next
    mdiMain.tcp.SendData "JOIN " & Item.Text & vbCrLf
End Sub
