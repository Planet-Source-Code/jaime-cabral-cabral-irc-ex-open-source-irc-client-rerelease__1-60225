Attribute VB_Name = "mdi_task"
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWNOACTIVE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9


'''''''''''''
'no titlebar
'no bar
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Sub ReleaseCapture Lib "user32" ()
