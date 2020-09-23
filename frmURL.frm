VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabral URL Catcher"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwURL 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRefresh_Click()
    lvwURL.ListItems.Clear
'    'load URL INFO
    If FileExists(App.Path & "\url.txt") Then
        CNumber = FreeFile
        'read nickname file
       Open App.Path & "\url.txt" For Input As #CNumber
        Do While Not (EOF(CNumber))
            Line Input #CNumber, iLINE
            If Trim(iLINE) <> "" Then
                word = Split(iLINE, Chr(1))
                Set mItem = frmURL.lvwURL.ListItems.Add(, , word(0))
                mItem.SubItems(1) = word(1)
                mItem.SubItems(2) = word(2)
            End If
        Loop
       Close #CNumber
    End If

End Sub

Private Sub Form_Load()
    Dim iLINE As String
    Dim word() As String

'load headers
    lvwURL.ColumnHeaders.Add , , "URL", 3000
    lvwURL.ColumnHeaders.Add , , "From"
    lvwURL.ColumnHeaders.Add , , "Date", 3500


'    'load URL INFO
    If FileExists(App.Path & "\url.txt") Then
        CNumber = FreeFile
        'read nickname file
       Open App.Path & "\url.txt" For Input As #CNumber
        Do While Not (EOF(CNumber))
            Line Input #CNumber, iLINE
            If Trim(iLINE) <> "" Then
                word = Split(iLINE, Chr(1))
                Set mItem = frmURL.lvwURL.ListItems.Add(, , word(0))
                mItem.SubItems(1) = word(1)
                mItem.SubItems(2) = word(2)
            End If
        Loop
       Close #CNumber
    End If
End Sub
