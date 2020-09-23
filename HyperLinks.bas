Attribute VB_Name = "HyperLinks"
'==================== Auto Detecter ====================
'Not Completely mine ;)
Option Explicit

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
    
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Const WM_NOTIFY = &H4E
Const EM_SETEVENTMASK = &H445
Const EM_GETEVENTMASK = &H43B
Const EM_GETTEXTRANGE = &H44B
Const EM_AUTOURLDETECT = &H45B
Const EN_LINK = &H70B

Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_SETCURSOR = &H20

Const CFE_LINK = &H20
Const ENM_LINK = &H4000000
Const GWL_WNDPROC = (-4)
Const SW_SHOW = 5

Dim lOldProc As Long    'Old windowproc
Dim hWndRTB As Long     'hWnd of RTB
Dim hWndParent As Long  'hWnd of parent window

Public Sub EnableURLDetect(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
    'Don't want to subclass twice!
    If lOldProc = 0 Then
        'Subclass!
        lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)
        SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
        SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
        hWndParent = hWndOwner
        hWndRTB = hWndTextbox
    End If
End Sub
Public Sub DisableURLDetect()
    If lOldProc Then
        SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
        'Reset the window procedure (stop the subclassing)
        SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
        'Set this to 0 so we can subclass again in future
        lOldProc = 0
    End If
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long
    
    'Which message?
    Select Case uMsg
    Case WM_NOTIFY
        'Copy the notification header into our structure from the pointer
        CopyMemory uHead, ByVal lParam, Len(uHead)
        
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
            
            'What kind of message?
            Select Case eLink.msg
            Case WM_LBUTTONDBLCLK 'Double clicked the link!
        eText.chrg.cpMin = eLink.chrg.cpMin
        eText.chrg.cpMax = eLink.chrg.cpMax
        eText.lpstrText = Space$(1024)
        lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
        sText = left$(eText.lpstrText, lLen)
        ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
                
            End Select
            
        End If
        
    End Select
    WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam)
End Function

