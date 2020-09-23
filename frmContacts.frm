VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContacts 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgContacts 
      Left            =   960
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":0E54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvContacts 
      Height          =   3835
      Left            =   30
      TabIndex        =   2
      Top             =   240
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6773
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgContacts"
      Appearance      =   0
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   840
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim i As Node
     trvContacts.Nodes.Add , , "Online", "Online", 1

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r%

        If Button = 1 Then
                ReleaseCapture
                r% = SendMessage(hWnd, WM_NCLBUTTONDOWN, 2, 0)
        End If

End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r%

        If Button = 1 Then
                ReleaseCapture
                r% = SendMessage(hWnd, WM_NCLBUTTONDOWN, 2, 0)
        End If

End Sub

