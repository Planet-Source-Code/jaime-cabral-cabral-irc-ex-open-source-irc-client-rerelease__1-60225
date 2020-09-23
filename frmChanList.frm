VERSION 5.00
Begin VB.Form frmChanList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabral Channel List"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmChanList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMAX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "9999"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Retrieve"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Number of users in channel:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmChanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdList_Click()
    txtSearch = Trim(txtSearch)
    If txtSearch = "" Then
        mdiMain.tcp.SendData "LIST >" & txtMIN & " <" & txtMAX & vbCrLf
        Unload Me
    Else
        txtSearch = "*" & txtSearch & "*"
        'MsgBox "LIST >" & txtMIN & " <" & txtMAX & " " & txtSearch
        mdiMain.tcp.SendData "LIST " & txtSearch & vbCrLf
        Unload Me
    End If
End Sub

Private Sub txtMAX_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc(vbCr)
            KeyAscii = 0
        Case 8
            'Backspace <BS>
        Case 46
            ' DOT .
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMIN_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc(vbCr)
            KeyAscii = 0
        Case 8
            'Backspace <BS>
        Case 46
            ' DOT .
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub
