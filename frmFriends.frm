VERSION 5.00
Begin VB.Form frmFriends 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabral Friends"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   Icon            =   "frmFriends.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   2550
   Begin VB.ListBox lstFriends 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmFriends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstFriends_DblClick()
    'load user set colors
    Dim colorX(0 To 15) As Long
    colorX(0) = vbWhite 'white
    colorX(1) = vbBlack 'black
    colorX(2) = RGB(0, 0, 140) 'dark blue
    colorX(3) = RGB(0, 140, 0) 'dark green
    colorX(4) = vbRed 'red
    colorX(5) = RGB(110, 65, 0) 'brown
    colorX(6) = RGB(140, 0, 140) 'purple
    colorX(7) = RGB(248, 146, 0) 'orange
    colorX(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    colorX(9) = vbGreen 'light green
    colorX(10) = RGB(0, 140, 140) 'dark blue green
    colorX(11) = RGB(0, 255, 255) 'light blue green
    colorX(12) = vbBlue 'light blue
    colorX(13) = vbMagenta 'magenta
    colorX(14) = RGB(140, 140, 140) 'grey
    colorX(15) = RGB(200, 200, 200) 'light grey





    Dim i As Integer
    Dim xlfound As Boolean
    xlfound = False
    Dim ChatWith As String
    ChatWith = lstFriends.List(lstFriends.ListIndex)
    ChatWith = Replace(ChatWith, "@", "")
    ChatWith = Replace(ChatWith, "+", "")
    ChatWith = Replace(ChatWith, "%", "")

    For i = 1 To 100
        If LCase(QueryName(i)) = LCase(ChatWith) Then
            xlfound = True
            Exit For
        End If
    Next i
    
    If xlfound = False Then
        For i = 1 To 100
            If QueryName(i) = "" Then
                Load Query(i)
                'set user colors
                Query(i).txtSend.BackColor = colorX(color.bgText)
                Query(i).txtSend.ForeColor = colorX(color.normal)
                Query(i).txtQuery.BackColor = colorX(color.bgText)
                '
                Query(i).Caption = ChatWith
                QueryName(i) = ChatWith
                Call AddTaskbar(ChatWith, 1)
                mdiMain.tvMain.Nodes.Add "mainChats", tvwChild, LCase(ChatWith), ChatWith, 3
                Exit For
            End If
        Next i
    End If
End Sub
