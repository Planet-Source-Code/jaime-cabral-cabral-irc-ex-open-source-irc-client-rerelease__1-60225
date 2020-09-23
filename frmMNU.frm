VERSION 5.00
Begin VB.Form frmMNU 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Visible         =   0   'False
   Begin VB.Menu mnuChannelNicks 
      Caption         =   "ChannelNicks"
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois?"
      End
      Begin VB.Menu mnuQuery 
         Caption         =   "Private Message"
      End
      Begin VB.Menu h0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControl 
         Caption         =   "Control"
         Begin VB.Menu mnuOp 
            Caption         =   "op"
         End
         Begin VB.Menu mnuDeop 
            Caption         =   "deop"
         End
         Begin VB.Menu mnuVoice 
            Caption         =   "voice"
         End
         Begin VB.Menu mnuDevoice 
            Caption         =   "devoice"
         End
      End
   End
End
Attribute VB_Name = "frmMNU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

