VERSION 5.00
Begin VB.Form frmListChannels 
   Caption         =   "Cabral Channel List"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   Icon            =   "frmListChannels.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   7245
   Begin VB.ListBox lstChannels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3150
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmListChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LB_SETTABSTOPS = &H192
Public Sub SetListTabStops(ListHandle As Long, _
    ParamArray ParmList() As Variant)
    Dim i As Long
    Dim ListTabs() As Long
    Dim NumColumns As Long

    ReDim ListTabs(UBound(ParmList))
    For i = 0 To UBound(ParmList)
        ListTabs(i) = ParmList(i)
    Next i
    NumColumns = UBound(ParmList) + 1

    Call SendMessage(ListHandle, LB_SETTABSTOPS, _
        NumColumns, ListTabs(0))
End Sub
Private Sub Form_Load()
Call SetListTabStops(lstChannels.hWnd, 0, 74, 100)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstChannels.Move lstChannels.Left, lstChannels.Top, Me.Width - 100, Me.Height - 400
End Sub

Private Sub lstChannels_DblClick()
    Dim CCC() As String
    CCC = Split(lstChannels.List(lstChannels.ListIndex), vbTab)
    mdiMain.tcp.SendData "JOIN " & CCC(0) & vbCrLf
End Sub
