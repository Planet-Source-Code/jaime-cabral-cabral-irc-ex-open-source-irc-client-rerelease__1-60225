VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cabral Script Editor"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   5895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdTEST 
      Caption         =   "Test"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   735
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    On Error Resume Next
    'mdiMain.script.AddCode txtCode
End Sub

Private Sub cmdTEST_Click()
    On Error Resume Next
    'mdiMain.script.ExecuteStatement txtCode
End Sub

