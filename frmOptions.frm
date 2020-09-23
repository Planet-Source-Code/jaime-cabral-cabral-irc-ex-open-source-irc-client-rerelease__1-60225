VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cabral Options"
   ClientHeight    =   6330
   ClientLeft      =   5670
   ClientTop       =   3600
   ClientWidth     =   7965
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7965
   Begin VB.CheckBox chkShowMe 
      Caption         =   "Show options on startup"
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
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   5880
      Width           =   975
   End
   Begin TabDlg.SSTab Tab 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   8
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Server"
      TabPicture(0)   =   "frmOptions.frx":1AFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line7(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbNetwork"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdNetworkAdd"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdNetworkDelete"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdServerAdd"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdServerEdit"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdServerDelete"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtServer"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtServerPort"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPassword"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lvwServers"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":1B16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNickname"
      Tab(1).Control(1)=   "cmdDefaultIdent"
      Tab(1).Control(2)=   "chkIdentShow"
      Tab(1).Control(3)=   "txtIdentUserID"
      Tab(1).Control(4)=   "txtIdentSystem"
      Tab(1).Control(5)=   "txtIdentPort"
      Tab(1).Control(6)=   "chkIdent"
      Tab(1).Control(7)=   "txtRealName"
      Tab(1).Control(8)=   "txtEmail"
      Tab(1).Control(9)=   "Label27"
      Tab(1).Control(10)=   "Line8"
      Tab(1).Control(11)=   "Line7(1)"
      Tab(1).Control(12)=   "Label26(2)"
      Tab(1).Control(13)=   "Label26(1)"
      Tab(1).Control(14)=   "Label26(0)"
      Tab(1).Control(15)=   "Label11(3)"
      Tab(1).Control(16)=   "Shape1(5)"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "IRC"
      TabPicture(2)   =   "frmOptions.frx":1B32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkWallops"
      Tab(2).Control(1)=   "chkServerMSG"
      Tab(2).Control(2)=   "chkInvisible"
      Tab(2).Control(3)=   "chkShowMOTD"
      Tab(2).Control(4)=   "Frame2"
      Tab(2).Control(5)=   "chkSkipMOTD"
      Tab(2).Control(6)=   "chkRejoin"
      Tab(2).Control(7)=   "chkAutoJoin"
      Tab(2).Control(8)=   "chkWhois"
      Tab(2).Control(9)=   "Label11(4)"
      Tab(2).Control(10)=   "Shape1(6)"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Contacts"
      TabPicture(3)   =   "frmOptions.frx":1B4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkNotifyList"
      Tab(3).Control(1)=   "chkWhoisNotify"
      Tab(3).Control(2)=   "cmdDelete"
      Tab(3).Control(3)=   "chkEnable"
      Tab(3).Control(4)=   "lstNotify"
      Tab(3).Control(5)=   "txtNotifyNickName"
      Tab(3).Control(6)=   "chkNotifyOnActive"
      Tab(3).Control(7)=   "cmdAdd"
      Tab(3).Control(8)=   "Label11(5)"
      Tab(3).Control(9)=   "Shape1(7)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Colors"
      TabPicture(4)   =   "frmOptions.frx":1B6A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picBGColor"
      Tab(4).Control(1)=   "Picture2"
      Tab(4).Control(2)=   "picColor(0)"
      Tab(4).Control(3)=   "picColor(1)"
      Tab(4).Control(4)=   "picColor(2)"
      Tab(4).Control(5)=   "picColor(3)"
      Tab(4).Control(6)=   "picColor(4)"
      Tab(4).Control(7)=   "picColor(5)"
      Tab(4).Control(8)=   "picColor(6)"
      Tab(4).Control(9)=   "picColor(7)"
      Tab(4).Control(10)=   "picColor(8)"
      Tab(4).Control(11)=   "picColor(9)"
      Tab(4).Control(12)=   "picColor(10)"
      Tab(4).Control(13)=   "picColor(11)"
      Tab(4).Control(14)=   "picColor(12)"
      Tab(4).Control(15)=   "picColor(13)"
      Tab(4).Control(16)=   "picColor(14)"
      Tab(4).Control(17)=   "picColor(15)"
      Tab(4).Control(18)=   "Label11(6)"
      Tab(4).Control(19)=   "Shape1(8)"
      Tab(4).ControlCount=   20
      TabCaption(5)   =   "Text Strings"
      TabPicture(5)   =   "frmOptions.frx":1B86
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtString"
      Tab(5).Control(1)=   "Frame1"
      Tab(5).Control(2)=   "lstString"
      Tab(5).Control(3)=   "Label11(7)"
      Tab(5).Control(4)=   "Shape1(9)"
      Tab(5).ControlCount=   5
      Begin VB.CheckBox chkNotifyList 
         Caption         =   "Popup Notify Window on Connect"
         Height          =   255
         Left            =   -70320
         TabIndex        =   120
         Top             =   4680
         Width           =   2895
      End
      Begin MSComctlLib.ListView lvwServers 
         Height          =   1695
         Left            =   600
         TabIndex        =   108
         Top             =   1920
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.TextBox txtString 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74760
         TabIndex        =   87
         Top             =   5160
         Width           =   7215
      End
      Begin VB.Frame Frame1 
         Height          =   3975
         Left            =   -72960
         TabIndex        =   80
         Top             =   1080
         Width           =   5415
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "$msg = message text"
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   1920
            Width           =   1485
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "$color = color code text. $color4 <- red text"
            Height          =   195
            Left            =   240
            TabIndex        =   109
            Top             =   1680
            Width           =   3030
         End
         Begin VB.Label Label23 
            Caption         =   "$kicked = person being kicked"
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "$reason = reason for action"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "$chan = channel name"
            Height          =   195
            Left            =   240
            TabIndex        =   84
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "$text = text message"
            Height          =   195
            Left            =   240
            TabIndex        =   83
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "$address = address"
            Height          =   195
            Left            =   240
            TabIndex        =   82
            Top             =   480
            Width           =   1380
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "$nick = nickname/kicker"
            Height          =   195
            Left            =   240
            TabIndex        =   81
            Top             =   240
            Width           =   1770
         End
      End
      Begin VB.ListBox lstString 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   -74880
         TabIndex        =   79
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkWhoisNotify 
         Caption         =   "Perform /WHOIS On Notify"
         Height          =   255
         Left            =   -71160
         TabIndex        =   78
         Top             =   3120
         Width           =   2895
      End
      Begin VB.PictureBox picBGColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFC0C0&
         ForeColor       =   &H00FFC0C0&
         Height          =   4215
         Left            =   -74280
         ScaleHeight     =   4185
         ScaleWidth      =   2625
         TabIndex        =   61
         Top             =   1200
         Width           =   2655
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normal Text"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background Color"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ctcp"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   75
            Top             =   600
            Width           =   315
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "notice"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Action"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invite Text"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   750
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Join"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   285
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kick"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   1800
            Width           =   315
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modes"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   69
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nick Changes"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   68
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notify"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   67
            Top             =   2520
            Width           =   405
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   66
            Top             =   2760
            Width           =   285
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quit"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   65
            Top             =   3000
            Width           =   285
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Topics"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   64
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Whois"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   63
            Top             =   3480
            Width           =   450
         End
         Begin VB.Label lblcolor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   62
            Top             =   3720
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -71520
         ScaleHeight     =   465
         ScaleWidth      =   2745
         TabIndex        =   59
         Top             =   1920
         Width           =   2775
         Begin VB.Label lblExample 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Normal Text"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   0
         Left            =   -71520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   58
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   1
         Left            =   -71160
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   2
         Left            =   -70800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   56
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   3
         Left            =   -70440
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   55
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   4
         Left            =   -70080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   54
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   5
         Left            =   -69720
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   53
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   6
         Left            =   -69360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   52
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   7
         Left            =   -69000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   51
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   8
         Left            =   -71520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   50
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   9
         Left            =   -71160
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   49
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   10
         Left            =   -70800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   48
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   11
         Left            =   -70440
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   47
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   12
         Left            =   -70080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   46
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   13
         Left            =   -69720
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   45
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   14
         Left            =   -69360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   44
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         Height          =   255
         Index           =   15
         Left            =   -69000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -70200
         TabIndex        =   42
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Enable Notify"
         Height          =   255
         Left            =   -74640
         TabIndex        =   41
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ListBox lstNotify 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   -74640
         TabIndex        =   40
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtNotifyNickName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71400
         TabIndex        =   39
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CheckBox chkNotifyOnActive 
         Caption         =   "Show Notify in Active Window"
         Height          =   255
         Left            =   -70320
         TabIndex        =   38
         Top             =   5040
         Width           =   2775
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -71160
         TabIndex        =   37
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtNickname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72960
         TabIndex        =   36
         Text            =   "Cabral"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtServerPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Text            =   "6667"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Text            =   "irc.dal.net"
         Top             =   4440
         Width           =   3375
      End
      Begin VB.CommandButton cmdServerDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   6240
         TabIndex        =   32
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdServerEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdServerAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4560
         TabIndex        =   30
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdNetworkDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   6240
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdNetworkAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   5400
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkWallops 
         Caption         =   "Op Messages"
         Height          =   255
         Left            =   -74640
         TabIndex        =   27
         Top             =   4680
         Width           =   3135
      End
      Begin VB.CheckBox chkServerMSG 
         Caption         =   "Server Messages"
         Height          =   255
         Left            =   -74640
         TabIndex        =   26
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CheckBox chkInvisible 
         Caption         =   "Invisible"
         Height          =   255
         Left            =   -74640
         TabIndex        =   25
         Top             =   4200
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CommandButton cmdDefaultIdent 
         Caption         =   "Default"
         Height          =   255
         Left            =   -72600
         TabIndex        =   24
         Top             =   4440
         Width           =   735
      End
      Begin VB.CheckBox chkIdentShow 
         Caption         =   "Show ident requests "
         Height          =   255
         Left            =   -71520
         TabIndex        =   23
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtIdentUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73800
         TabIndex        =   22
         Text            =   "cabral"
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtIdentSystem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73800
         TabIndex        =   21
         Text            =   "UNIX"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox txtIdentPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -73800
         TabIndex        =   20
         Text            =   "113"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox chkIdent 
         Caption         =   "Enable Ident server"
         Height          =   255
         Left            =   -71520
         TabIndex        =   19
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtRealName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72960
         TabIndex        =   18
         Text            =   "Cabral IRC Client"
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -72960
         TabIndex        =   17
         Text            =   "cabral"
         Top             =   1920
         Width           =   3135
      End
      Begin VB.ComboBox cmbNetwork 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   16
         Text            =   "AllGroups"
         Top             =   1080
         Width           =   5895
      End
      Begin VB.CheckBox chkShowMOTD 
         Caption         =   "Show MOTD on seperate window"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Show in channel:"
         Height          =   1575
         Left            =   -71040
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
         Begin VB.CheckBox chkShowKicks 
            Caption         =   "Kicks"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkShowTopics 
            Caption         =   "Topics"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkShowModes 
            Caption         =   "Modes"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkShowJoinPart 
            Caption         =   "Parts/Joins"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkShowQuits 
            Caption         =   "Quits"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkSkipMOTD 
         Caption         =   "Skip Message of the Day on connect"
         Height          =   255
         Left            =   -74640
         TabIndex        =   8
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CheckBox chkRejoin 
         Caption         =   "rejoin channel when kicked"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkAutoJoin 
         Caption         =   "Auto join channel on invite"
         Height          =   255
         Left            =   -74640
         TabIndex        =   6
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox chkWhois 
         Caption         =   "Whois on query"
         Height          =   255
         Left            =   -74640
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "Identd Setup:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   119
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   -67560
         X2              =   -74880
         Y1              =   2990
         Y2              =   2990
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   -67560
         X2              =   -74880
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   118
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Name:"
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   117
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Nickname:"
         Height          =   255
         Index           =   0
         Left            =   -74280
         TabIndex        =   116
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Text Display"
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
         Index           =   7
         Left            =   -74640
         TabIndex        =   115
         Top             =   720
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   9
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Color Options"
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
         Index           =   6
         Left            =   -74640
         TabIndex        =   114
         Top             =   720
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   8
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Friend List"
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
         Index           =   5
         Left            =   -74640
         TabIndex        =   113
         Top             =   720
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   7
         Left            =   -74750
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "IRC Options"
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
         Index           =   4
         Left            =   -74640
         TabIndex        =   112
         Top             =   720
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   6
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "General Options"
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
         Index           =   3
         Left            =   -74640
         TabIndex        =   111
         Top             =   720
         Width           =   1560
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   5
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "Server Text Settings"
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
         Index           =   2
         Left            =   -74640
         TabIndex        =   107
         Top             =   720
         Width           =   2010
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   4
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Text:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   106
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IRC Colors"
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
         Left            =   -74640
         TabIndex        =   105
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact List"
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
         Left            =   -74640
         TabIndex        =   104
         Top             =   720
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         Height          =   195
         Index           =   1
         Left            =   -72360
         TabIndex        =   103
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   720
         TabIndex        =   102
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Left            =   1080
         TabIndex        =   101
         Top             =   4800
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Left            =   960
         TabIndex        =   100
         Top             =   4440
         Width           =   510
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   7320
         X2              =   1680
         Y1              =   4215
         Y2              =   4215
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         X1              =   1680
         X2              =   7320
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Default Server"
         Height          =   195
         Left            =   480
         TabIndex        =   99
         Top             =   4080
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "IRC Servers"
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
         Index           =   1
         Left            =   360
         TabIndex        =   98
         Top             =   720
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000011&
         BackStyle       =   0  'Transparent
         Caption         =   "General IRC Information"
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   97
         Top             =   720
         Width           =   2385
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000014&
         X1              =   -73200
         X2              =   -67560
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -73200
         X2              =   -67560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Personal Information"
         Height          =   195
         Left            =   -74760
         TabIndex        =   96
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         X1              =   -73800
         X2              =   -67560
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   -73800
         X2              =   -67560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ident Server"
         Height          =   195
         Left            =   -74760
         TabIndex        =   95
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   94
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "System:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   93
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   92
         Top             =   3960
         Width           =   330
      End
      Begin VB.Label Label4 
         Caption         =   "Real Name/Phrase:"
         Height          =   255
         Left            =   -74520
         TabIndex        =   91
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   90
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nickname:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   89
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Network:"
         Height          =   195
         Left            =   480
         TabIndex        =   88
         Top             =   1080
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H80000011&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   3
         Left            =   -74760
         Top             =   600
         Width           =   7215
      End
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   120
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbServers_Change()

End Sub



Private Sub cmbNetwork_Click()
    'cmbNetwork.Text
    'LOAD servers
    Dim word() As String
    Dim varX As String
    Dim i As Integer
    iniFilename = "servers.ini"
    lvwServers.ListItems.Clear
    For i = 0 To 500
        varX = i
        OutPut = ReadINI(cmbNetwork.Text, varX)
        'If OutPut = "" Then Exit For
        If OutPut <> "" Then
            word = Split(OutPut, "|")
            Set mItem = lvwServers.ListItems.Add(, , word(1))
            mItem.SubItems(1) = word(0)
            mItem.SubItems(2) = word(2)
        End If
    Next i

    
End Sub

Private Sub cmdAddNickname_Click()
    frmAddNickName.SHOW 0, mdiMain
End Sub

Private Sub cmdAddServer_Click()
    frmAddServer.SHOW 0, mdiMain
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer
    If Trim(txtNickname) <> "" Then
        i = CheckListbox(lstNotify, Trim(txtNotifyNickName))
        If i = 0 Then
            lstNotify.AddItem Trim(txtNotifyNickName)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()

    Dim xServer As String
    Dim xPort As String
    Dim i As Integer
    Dim PortFound As Boolean
    
    '''''''''''''''''''''''''''
    'On Connected Modes
    iModes.i = chkInvisible.Value
    iModes.s = chkServerMSG.Value
    iModes.w = chkWallops.Value
    'they will be sent to server when
    'motd is first started
    'goto numeric code 375
    '''''''''''''''''''''''''''

    xServer = Trim(txtServer)
    xPort = Val(Trim(txtServerPort))
    If xServer = "" Then xServer = "127.0.0.1"
    If xPort = "" Then xPort = "6667"

    email = txtEmail
    RealName = txtRealName
    nickname = txtNickname
    With mdiMain.tcp
        If .State <> 0 Then
            .Close
        End If
        .Connect xServer, xPort
        DoColor frmStatus.txtStatus, "2*** Attempting to connect to " & xServer & ":" & xPort & vbCrLf & "-" & vbCrLf
    End With
    Call cmdOK_Click
End Sub

Private Sub cmdDefaultIdent_Click()
    txtIdentPort = "113"
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    For i = 0 To lstNotify.ListCount - 1
        If lstNotify.List(lstNotify.ListIndex) = lstNotify.List(i) Then
            lstNotify.RemoveItem i
            Exit For
        End If
    Next i
End Sub

Private Sub cmdNetworkAdd_Click()
    frmAddNetwork.SHOW 1, frmOptions
End Sub

Private Sub cmdNetworkDelete_Click()

    Dim delNet As String
    Dim iNet As String
    iNet = cmbNetwork.Text
    
    If LCase(iNet) = "allgroups" Then Exit Sub
    delNet = WriteINI(cmbNetwork.Text, vbNullString, vbNullString)
    
    Dim i As Integer
    For i = 0 To cmbNetwork.ListCount - 1
        If LCase(cmbNetwork.List(i)) = LCase(cmbNetwork.Text) Then
            cmbNetwork.RemoveItem i
            Exit For
        End If
    Next i
    
    Dim varX As String
    For i = 0 To 500
        varX = i
        delNet = ReadINI("AllGroups", varX)
        If LCase(delNet) = LCase(iNet) Then
            delNet = WriteINI("AllGroups", varX, vbNullString)
            delNet = WriteINI(iNet, vbNullString, vbNullString)
            Exit For
        End If
    Next i
    
    cmbNetwork.Text = "AllGroups"
End Sub

Private Sub cmdOK_Click()
    Dim X As Integer
    
    'write contact info
        'save nicknames to file
    'and add to array
    For X = 1 To 100
        notify(X) = ""
    Next X
    notifylist = ""
    Open App.Path & "\notify.txt" For Output As #1
    For X = 0 To lstNotify.ListCount - 1
        Print #1, lstNotify.List(X)
        notify(X + 1) = lstNotify.List(X)
        notifylist = notifylist & notify(X + 1) & " "
    Next X
    Close #1
    
    'old nickname notify
    If chkEnable = 1 Then
        mdiMain.ISON.Enabled = True
    Else
        mdiMain.ISON.Enabled = False
    End If

    
    
    
    Dim i As String
    iniFilename = "irc.ini"
    'save ident options
    i = WriteINI("IDENT", "IDENT", chkIdent.Value)
    i = WriteINI("IDENT", "SHOW", chkIdentShow.Value)
    i = WriteINI("IDENT", "PORT", Trim(txtIdentPort))
    i = WriteINI("IDENT", "USERID", Trim(txtIdentUserID))
    IdentUserID = Trim(txtIdentUserID)
    i = WriteINI("IDENT", "SYSTEM", Trim(txtIdentSystem))
    
    'save server options
    i = WriteINI("INFO", "SERVER", txtServer)
    i = WriteINI("INFO", "PORT", txtServerPort)
    i = WriteINI("INFO", "NICKNAME", txtNickname.Text)
    i = WriteINI("INFO", "USERNAME", txtEmail)
    i = WriteINI("INFO", "REALNAME", txtRealName)
    
    'SAVE irc options
    i = WriteINI("SHOW", "QUITS", chkShowQuits.Value)
    iShow.quits = chkShowQuits.Value
    i = WriteINI("SHOW", "JOINPART", chkShowJoinPart.Value)
    iShow.joinpart = chkShowJoinPart.Value
    i = WriteINI("SHOW", "MODES", chkShowModes.Value)
    iShow.modes = chkShowModes.Value
    i = WriteINI("SHOW", "TOPICS", chkShowTopics.Value)
    iShow.topics = chkShowTopics.Value
    i = WriteINI("SHOW", "KICKS", chkShowKicks.Value)
    iShow.kicks = chkShowKicks.Value
    i = WriteINI("IRC", "WHOIS", chkWhois.Value)
    i = WriteINI("IRC", "AUTOJOIN", chkAutoJoin.Value)
    i = WriteINI("IRC", "REJOIN", chkRejoin.Value)
    i = WriteINI("IRC", "SKIPMOTD", chkSkipMOTD.Value)
    i = WriteINI("IRC", "WHOISNOTIFY", chkWhoisNotify.Value)
    iShow.whoisnotify = chkWhoisNotify.Value
    i = WriteINI("IRC", "SHOWMOTD", chkShowMOTD.Value)
    iShow.motd = chkShowMOTD.Value
    'i = WriteINI("IRC", "SHOWADDRESS", chkShowAddress.Value)
    'iShow.address = chkShowAddress.Value

    'Show options at startup?
    i = WriteINI("SHOW", "OPTIONS", chkShowMe.Value)
    
    'Is the notify on?
    i = WriteINI("ISON", "ISON", chkEnable.Value)
    i = WriteINI("ISON", "LIST", chkNotifyList.Value)
    
    'Custom Text Strings
    i = WriteINI("CUSTOM", "JOIN", strCustom.join)
    i = WriteINI("CUSTOM", "PART", strCustom.part)
    i = WriteINI("CUSTOM", "KICK", strCustom.kick)
    i = WriteINI("CUSTOM", "QUIT", strCustom.quit)
    i = WriteINI("CUSTOM", "PM", strCustom.pm)
    
    
    'Let's write color information to INI file
    Dim srColor As String
    Dim TwoDigitColor As String
    For X = 0 To 15
        If Len(lblcolor(X).Tag) = 1 Then lblcolor(X).Tag = "0" & lblcolor(X).Tag
        srColor = srColor & X & ":" & lblcolor(X).Tag & " "
        Select Case X
            Case 0
                color.bgText = lblcolor(X).Tag
            Case 1
                color.normal = lblcolor(X).Tag
            Case 2
                color.ctcp = lblcolor(X).Tag
            Case 3
                color.notice = lblcolor(X).Tag
            Case 4
                color.action = lblcolor(X).Tag
            Case 5
                color.invite = lblcolor(X).Tag
            Case 6
                color.join = lblcolor(X).Tag
            Case 7
                color.kick = lblcolor(X).Tag
            Case 8
                color.mode = lblcolor(X).Tag
            Case 9
                color.nick = lblcolor(X).Tag
            Case 10
                color.notify = lblcolor(X).Tag
            Case 11
                color.part = lblcolor(X).Tag
            Case 12
                color.quit = lblcolor(X).Tag
            Case 13
                color.topic = lblcolor(X).Tag
            Case 14
                color.whois = lblcolor(X).Tag
            Case 15
                color.server = lblcolor(X).Tag
        End Select
    Next X
    i = WriteINI("IRC", "COLORS", srColor)

    
    
    If chkIdent Then
        On Error GoTo InUse
        mdiMain.ident(0).Close
        If mdiMain.ident(0).State <> sckListening Then
            mdiMain.ident(0).LocalPort = Val(txtIdentPort)
            'mdiMain.ident(0).Bind Val(txtIdentPort), mdiMain.ident(0).LocalIP
            mdiMain.ident(0).Listen
        End If
    End If
    
    'OK...unload form
    Unload Me
'Identd socket is used in another program - conflict
InUse:
    If Err.Number = 10048 Then
        MsgBox "Another program is using port " & mdiMain.ident(0).LocalPort & "." & vbCrLf & "IdentD will be disabled." & vbCrLf & "You will need to close the other program and" & vbCrLf & "reopen to use Cabral's Ident server.", vbOKOnly, "Cabral IRC IdentD problem"
        Resume Next
    End If

End Sub

Private Sub cmdServerAdd_Click()
    frmAddServer.SHOW 1, Me
End Sub

Private Sub cmdServerDelete_Click()
    On Error Resume Next
    lvwServers.ListItems.Remove lvwServers.SelectedItem.Index
    If lvwServers.ListItems.Count = 0 Then
        Dim i As Integer
        i = MsgBox("Would you like to delete the Network?", vbYesNo + vbExclamation, "Network Empty")
        '6 = yes delete
        If i = 6 Then
            Call cmdNetworkDelete_Click
        End If
    End If
End Sub

Private Sub cmdServerEdit_Click()
    frmEditServer.SHOW 1, Me
End Sub


Private Sub Form_Load()
    On Error Resume Next
    'usually error occurs due to a lack of files - but they'll be created and the the problem goes away
    Dim i As Integer
    Dim varX As String
    Dim OutPut As String

    'LOAD NETWORKS
    iniFilename = "servers.ini"
    For i = 0 To 500
        varX = i
        OutPut = ReadINI("ALLGROUPS", varX)
        'If OutPut = "" Then Exit For
        If OutPut <> "" Then
            cmbNetwork.AddItem OutPut
        End If
    Next i

    'load Contact information
    Dim users As String
    Dim iNotifyString As String
    If FileExists(App.Path & "\notify.txt") Then
        CNumber = FreeFile
        'read nickname file
        Open App.Path & "\notify.txt" For Input As #CNumber
        Do While Not (EOF(CNumber))
            Line Input #CNumber, iNotifyString
            If Trim(iNotifyString) <> "" Then
                lstNotify.AddItem iNotifyString
            End If
        Loop
        Close #CNumber
    End If




    'LOAD ident options
    iniFilename = "irc.ini"
    chkIdent.Value = ReadINI("IDENT", "IDENT")
    chkIdentShow.Value = ReadINI("IDENT", "SHOW")
    txtIdentPort = ReadINI("IDENT", "PORT")
    txtIdentUserID = ReadINI("IDENT", "USERID")
    txtIdentSystem = ReadINI("IDENT", "SYSTEM")

    'LOAD server options
    txtServer = ReadINI("INFO", "SERVER")
    txtServerPort = ReadINI("INFO", "PORT")
    txtNickname.Text = ReadINI("INFO", "NICKNAME")
    txtEmail = ReadINI("INFO", "USERNAME")
    txtRealName = ReadINI("INFO", "REALNAME")

    'LOAD irc options
    chkShowQuits.Value = ReadINI("SHOW", "QUITS")
    chkShowJoinPart.Value = ReadINI("SHOW", "JOINPART")
    chkShowModes.Value = ReadINI("SHOW", "MODES")
    chkShowTopics.Value = ReadINI("SHOW", "TOPICS")
    chkShowKicks.Value = ReadINI("SHOW", "KICKS")
    chkWhois.Value = ReadINI("IRC", "WHOIS")
    chkAutoJoin.Value = ReadINI("IRC", "AUTOJOIN")
    chkRejoin.Value = ReadINI("IRC", "REJOIN")
    chkSkipMOTD.Value = ReadINI("IRC", "SKIPMOTD")
    chkShowMOTD.Value = ReadINI("IRC", "SHOWMOTD")
    chkShowAddress.Value = ReadINI("IRC", "SHOWADDRESS")
    chkWhoisNotify.Value = ReadINI("IRC", "WHOISNOTIFY")
    'LOAD Show Form at startup
    chkShowMe.Value = ReadINI("SHOW", "OPTIONS")
    'LOAD contact info
    chkEnable.Value = ReadINI("ISON", "ISON")
    chkNotifyList.Value = ReadINI("ISON", "LIST")
    
    'toolbar
    'mdiMain.Toolbar.Buttons(1).Value = tbrPressed
    
    'Server List view headers
    lvwServers.ColumnHeaders.Add , , "Description", 2000
    lvwServers.ColumnHeaders.Add , , "Server"
    lvwServers.ColumnHeaders.Add , , "Ports", 800
    'Set mItem = lvwServers.ListItems.Add(, , "Description")
    'mItem.SubItems(1) = "irc.server.com"
    
    
    'LOAD colors
    Dim color(0 To 15) As Long
    color(0) = vbWhite
    color(1) = vbBlack
    color(2) = RGB(42, 42, 87)
    color(3) = RGB(33, 112, 33)
    color(4) = vbRed
    color(5) = RGB(109, 50, 50)
    color(6) = RGB(119, 33, 119)
    color(7) = RGB(252, 127, 0)
    color(8) = RGB(195, 195, 56)
    color(9) = RGB(0, 252, 0)
    color(10) = RGB(89, 167, 179)
    color(11) = RGB(0, 255, 255)
    color(12) = vbBlue
    color(13) = RGB(255, 0, 255)
    color(14) = RGB(127, 127, 127)
    color(15) = RGB(210, 210, 210)
    
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
    
    
    'Fill Text Settings
    lstString.AddItem "Join"
    lstString.AddItem "Part"
    lstString.AddItem "Quit"
    lstString.AddItem "Kick"
    lstString.AddItem "PM"
    
    strCustom.join = ReadINI("CUSTOM", "JOIN")
    strCustom.part = ReadINI("CUSTOM", "PART")
    strCustom.kick = ReadINI("CUSTOM", "KICK")
    strCustom.quit = ReadINI("CUSTOM", "QUIT")
    strCustom.pm = ReadINI("CUSTOM", "PM")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'mdiMain.Toolbar.Buttons(1).Value = tbrUnpressed
End Sub

Private Sub lstString_Click()
    Select Case LCase(lstString.List(lstString.ListIndex))
        Case "join"
            txtString.Text = strCustom.join
        Case "part"
            txtString.Text = strCustom.part
        Case "kick"
            txtString.Text = strCustom.kick
        Case "quit"
            txtString.Text = strCustom.quit
        Case "pm"
            txtString.Text = strCustom.pm
    End Select
End Sub

Private Sub lvwServers_Click()
    On Error Resume Next
    'MsgBox lvwServers.SelectedItem.Index 'index in listbox
    'MsgBox lvwServers.SelectedItem.SubItems(1) 'server
    'MsgBox lvwServers.SelectedItem.SubItems(2) 'ports
    Me.txtServer = lvwServers.SelectedItem.SubItems(1)
    Me.txtServerPort = lvwServers.SelectedItem.SubItems(2)
    
    'loop through lvwbox
    'For i = 1 To lvwServers.ListItems.Count
    '    MsgBox lvwServers.ListItems(i).Text
    '    MsgBox lvwServers.ListItems(i).SubItems(1)
    '    MsgBox lvwServers.ListItems(i).SubItems(2)
    'Next i
End Sub

Private Sub picColor_Click(Index As Integer)
    Dim color(0 To 15) As Long
    color(0) = vbWhite
    color(1) = vbBlack
    color(2) = RGB(42, 42, 87)
    color(3) = RGB(33, 112, 33)
    color(4) = vbRed
    color(5) = RGB(109, 50, 50)
    color(6) = RGB(119, 33, 119)
    color(7) = RGB(252, 127, 0)
    color(8) = RGB(195, 195, 56)
    color(9) = RGB(0, 252, 0)
    color(10) = RGB(89, 167, 179)
    color(11) = RGB(0, 255, 255)
    color(12) = vbBlue
    color(13) = RGB(255, 0, 255)
    color(14) = RGB(127, 127, 127)
    color(15) = RGB(210, 210, 210)





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
Private Sub lblcolor_Click(Index As Integer)
    lblExample.Caption = lblcolor(Index).Caption
End Sub

Private Sub picBGColor_Click()
    lblExample.Caption = "Background Color"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtString_Change()
    Select Case LCase(lstString.List(lstString.ListIndex))
        Case "join"
            strCustom.join = txtString.Text
        Case "part"
            strCustom.part = txtString.Text
        Case "kick"
            strCustom.kick = txtString.Text
        Case "quit"
            strCustom.quit = txtString.Text
        Case "pm"
            strCustom.pm = txtString.Text
    End Select
End Sub
