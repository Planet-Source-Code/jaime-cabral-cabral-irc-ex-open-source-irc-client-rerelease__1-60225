Attribute VB_Name = "scroll_rtf"
Option Explicit

' Rtbscrol sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 2001 by Matt Hart
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/free/rtbscrol.htm"
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400
Public Const WM_DESTROY = &H2
Public Const WM_PARENTNOTIFY = &H210
Public Const EM_SCROLLCARET = &HB7
Public Const EM_REPLACESEL = &HC2
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const WM_SETFOCUS = &H7
Public Const WM_VSCROLL = &H115
Public Const SB_VERT = 1

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public origWndProc As Long, bAllowScroll As Boolean

Public Sub SetHook(hwnd, bSet As Boolean)
    If bSet Then
        origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
    ElseIf origWndProc Then
        Dim lRet As Long
        lRet = SetWindowLong(hwnd, GWL_WNDPROC, origWndProc)
        DeleteSetting "RtbScrol"
        origWndProc = 0
    End If
End Sub

Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static lOld As Long, bNotBottom As Boolean
    If Msg <> WM_DESTROY And lOld = 0 And origWndProc <> 0 Then
        lOld = origWndProc
        SaveSetting "RtbScrol", "WndProc", "WndProc", Str(lOld)
    ElseIf origWndProc = 0 Then
        GoTo DestroyIt
    End If
    Dim k As Long, lNumFiles As Long, l As Long
    Select Case Msg
        Case WM_VSCROLL, EM_REPLACESEL
            bNotBottom = False
            Dim S As SCROLLINFO
            S.cbSize = Len(S)
            S.fMask = SIF_ALL
            If GetScrollInfo(hwnd, SB_VERT, S) Then
                If S.nMax Then
                    If S.nPos < S.nMax - (S.nPage - 1) Then bNotBottom = True
                End If
            End If
        Case WM_SETFOCUS, EM_SCROLLCARET
            If Not bAllowScroll Or bNotBottom Then
                AppWndProc = True
                Exit Function
            End If
        Case EM_SETCHARFORMAT
            If Not bAllowScroll Or bNotBottom Then
                LockWindowUpdate hwnd
                AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
                LockWindowUpdate 0
                Exit Function
            End If
        Case WM_PARENTNOTIFY
            If (wParam And &HFF) = WM_DESTROY Then GoTo DestroyIt
        Case WM_DESTROY
DestroyIt:
            l = Val(GetSetting("RtbScrol", "WndProc", "WndProc"))
            SetWindowLong hwnd, GWL_WNDPROC, l
            AppWndProc = CallWindowProc(l, hwnd, Msg, wParam, lParam)
            origWndProc = 0
            DeleteSetting "RtbScrol", "WndProc", "WndProc"
            Exit Function
    End Select
    AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
End Function

