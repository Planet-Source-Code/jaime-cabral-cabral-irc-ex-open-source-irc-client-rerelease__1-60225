VERSION 5.00
Begin VB.Form frmColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cabral Color Changes"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Accept"
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim srColor As String
    Dim TwoDigitColor As String
    For i = 0 To 15
        If Len(lblcolor(i).Tag) = 1 Then lblcolor(i).Tag = "0" & lblcolor(i).Tag
        srColor = srColor & i & ":" & lblcolor(i).Tag & " "
        Select Case i
            Case 0
                color.bgText = lblcolor(i).Tag
            Case 1
                color.normal = lblcolor(i).Tag
            Case 2
                color.ctcp = lblcolor(i).Tag
            Case 3
                color.notice = lblcolor(i).Tag
            Case 4
                color.action = lblcolor(i).Tag
            Case 5
                color.invite = lblcolor(i).Tag
            Case 6
                color.join = lblcolor(i).Tag
            Case 7
                color.kick = lblcolor(i).Tag
            Case 8
                color.mode = lblcolor(i).Tag
            Case 9
                color.nick = lblcolor(i).Tag
            Case 10
                color.notify = lblcolor(i).Tag
            Case 11
                color.part = lblcolor(i).Tag
            Case 12
                color.quit = lblcolor(i).Tag
            Case 13
                color.topic = lblcolor(i).Tag
            Case 14
                color.whois = lblcolor(i).Tag
            Case 15
                color.server = lblcolor(i).Tag
        End Select
    Next i
    i = WriteINI("IRC", "COLORS", srColor)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim color(0 To 15) As Long
    color(0) = vbWhite 'white
    color(1) = vbBlack 'black
    color(2) = RGB(0, 0, 140) 'dark blue
    color(3) = RGB(0, 140, 0) 'dark green
    color(4) = vbRed 'red
    color(5) = RGB(110, 65, 0) 'brown
    color(6) = RGB(140, 0, 140) 'purple
    color(7) = RGB(248, 146, 0) 'orange
    color(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    color(9) = vbGreen 'light green
    color(10) = RGB(0, 140, 140) 'dark blue green
    color(11) = RGB(0, 255, 255) 'light blue green
    color(12) = vbBlue 'light blue
    color(13) = vbMagenta 'magenta
    color(14) = RGB(140, 140, 140) 'grey
    color(15) = RGB(200, 200, 200) 'light grey
    
    Dim i As Integer
    For i = o To 15
        picColor(i).BackColor = color(i)
    Next i
    
    srColor = ReadINI("IRC", "COLORS")
    Dim getColors() As String
    Dim TagColor() As String
    getColors = Split(srColor, " ")
    For i = 0 To UBound(getColors)
        TagColor = Split(getColors(i), ":")
        lblcolor(TagColor(0)).Tag = TagColor(1)
        lblcolor(TagColor(0)).ForeColor = color(TagColor(1))
        If i = 0 Then
            picBGColor.BackColor = color(TagColor(1))
        End If
    Next i
    lblcolor(0).ForeColor = lblcolor(1).ForeColor
End Sub

Private Sub lblcolor_Click(Index As Integer)
    lblExample.Caption = lblcolor(Index).Caption
End Sub

Private Sub picBGColor_Click()
    lblExample.Caption = "Background Color"
End Sub

Private Sub picColor_Click(Index As Integer)
    Dim color(0 To 15) As Long
    color(0) = vbWhite 'white
    color(1) = vbBlack 'black
    color(2) = RGB(0, 0, 140) 'dark blue
    color(3) = RGB(0, 140, 0) 'dark green
    color(4) = vbRed 'red
    color(5) = RGB(110, 65, 0) 'brown
    color(6) = RGB(140, 0, 140) 'purple
    color(7) = RGB(248, 146, 0) 'orange
    color(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    color(9) = vbGreen 'light green
    color(10) = RGB(0, 140, 140) 'dark blue green
    color(11) = RGB(0, 255, 255) 'light blue green
    color(12) = vbBlue 'light blue
    color(13) = vbMagenta 'magenta
    color(14) = RGB(140, 140, 140) 'grey
    color(15) = RGB(200, 200, 200) 'light grey





    For i = 0 To 15
        If LCase(lblExample.Caption) = LCase(lblcolor(i).Caption) Then
            lblcolor(i).ForeColor = color(Index)
            lblExample.ForeColor = color(Index)
            If LCase(lblExample.Caption) = "background color" Then
                lblExample.BackColor = color(Index)
                lblExample.ForeColor = lblcolor(1).ForeColor
                picBGColor.BackColor = color(Index)
            End If
            lblcolor(i).Tag = Index
        End If
    Next i
End Sub
