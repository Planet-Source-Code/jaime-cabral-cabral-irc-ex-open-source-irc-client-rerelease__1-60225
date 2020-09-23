VERSION 5.00
Begin VB.Form frmAddServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Server"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmAddServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Port Range:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Server Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server Description:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmAddServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Set mItem = frmOptions.lvwServers.ListItems.Add(, , txtDescription)
    mItem.SubItems(1) = txtName
    mItem.SubItems(2) = txtPort
    
    
    'loop through lvwbox
    'For i = 1 To lvwServers.ListItems.Count
    '    MsgBox lvwServers.ListItems(i).Text 'description
    '    MsgBox lvwServers.ListItems(i).SubItems(1) 'server
    '    MsgBox lvwServers.ListItems(i).SubItems(2) 'ports
    'Next i
    
    
    Dim word() As String
    Dim varX As String
    Dim i As Integer
    Dim NetFound As Boolean
    NetFound = False
    iniFilename = "servers.ini"
    'clear sections
    OutPut = WriteINI(frmOptions.cmbNetwork.Text, vbNullString, vbNullString)
    For i = 0 To 500
        varX = i
        OutPut = ReadINI("AllGroups", varX)
        If LCase(OutPut) = LCase(frmOptions.cmbNetwork.Text) Then
            MsgBox "Network already exsists", vbOKOnly + vbInformation
            NetFound = True
        End If
    Next i
    If NetFound = False Then
        For i = 0 To 500
            varX = i
            OutPut = ReadINI("AllGroups", varX)
            If OutPut = "" Then
                MsgBox "Let's add the server at " & varX, vbOKOnly + vbInformation
                OutPut = WriteINI("AllGroups", varX, frmOptions.cmbNetwork.Text)
                Exit For
            End If
        Next i
    End If
    'readd servers to section
    'ini = server, description, port
    For i = 1 To frmOptions.lvwServers.ListItems.Count
        varX = i - 1
        OutPut = WriteINI(frmOptions.cmbNetwork.Text, varX, frmOptions.lvwServers.ListItems(i).SubItems(1) & "|" & frmOptions.lvwServers.ListItems(i).Text & "|" & frmOptions.lvwServers.ListItems(i).SubItems(2))
    Next i
    
    
    Unload Me
End Sub
