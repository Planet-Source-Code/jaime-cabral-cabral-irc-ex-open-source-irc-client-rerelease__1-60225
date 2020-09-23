Attribute VB_Name = "ColorRTF"
Option Explicit
'back color text
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_USER = &H400
Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)


Public charf As CHARFORMAT2
Public Const LF_FACESIZE = 32
Public Const CFM_BACKCOLOR = &H4000000
Public Type CHARFORMAT2
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
    
    ' Additional stuff supported by RICHEDIT20
    wWeight As Integer            ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer           ' /* Amount to space between letters  */
    crBackColor As Long        ' /* Background color                 */
    lLCID As Long               ' /* Locale ID                        */
    dwReserved As Long         ' /* Reserved. Must be 0              */
    sStyle As Integer            ' /* Style handle                     */
    wKerning As Integer            ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte     ' /* Underline type                   */
    bAnimation As Byte         ' /* Animated text like marching ants */
    bRevAuthor As Byte         ' /* Revision author index            */
    bReserved1 As Byte
End Type

Public color As UserColor
Public Type UserColor
    normal As Integer
    bgText As Integer
    ctcp As Integer
    notice As Integer
    action As Integer
    invite As Integer
    join As Integer
    kick As Integer
    mode As Integer
    nick As Integer
    notify As Integer
    part As Integer
    quit As Integer
    topic As Integer
    whois As Integer
    server As Integer
End Type




Public Sub DoColor(RTF As RichTextBox, strData As String)
    On Error GoTo PassColorCode
    
    'lock rtf
    'LockWindowUpdate RTF.hWnd

    'chr(2) = bold
    'chr(3) = color
    'Chr(31) = Underline

    'haven't figured what this one does, it's pretty useless
    'so we'll delete this irc text code so it doesn't show up
    strData = Replace(strData, "", "")
    
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

    
    Dim strColor() As String
    Dim strBold() As String
    Dim strUnderline() As String
    Dim strTempColor As String
    Dim ColorCode As String
    Dim ColorCode2 As String
    Dim i As Integer
    Dim X As Integer
    Dim z As Integer
    Dim iColor As String
    Dim COLOR2 As Boolean
    Dim ret As Long
    Dim colorX As UserColor
    
    'default
    RTF.SelColor = color(colorX.normal)
    
    i = InStr(strData, Chr(3))
    If i Then
        strColor = Split(strData, Chr(3))
        For X = 0 To UBound(strColor)
            COLOR2 = False
            'lets get the first two color codes
            strColor(X) = Replace(strColor(X), Chr(3), "")
            iColor = Mid(strColor(X), 1, 5)
            For i = 1 To Len(iColor)
                If Mid(iColor, i, 1) = "," Then
                    ColorCode = Mid(Val(iColor), 1, i - 1)
                    COLOR2 = True
                    Exit For
                End If
            Next i
            If COLOR2 = False Then
                If IsNumeric(Mid(iColor, 1, 2)) Then
                    ColorCode = Mid(iColor, 1, 2)
                    If Right(ColorCode, 1) = "-" Or Right(ColorCode, 1) = "+" Then
                        ColorCode = Mid(ColorCode, 1, Len(ColorCode) - 1)
                    End If
                Else
                    If IsNumeric(Mid(iColor, 1, 1)) Then
                        ColorCode = CInt(Mid(iColor, 1, 1))
                    Else
                        ColorCode = ""
                    End If
                End If
            Else 'second color
                strTempColor = iColor
                strTempColor = Mid(iColor, Len(ColorCode) + 1)
                iColor = ""
                If Mid(strTempColor, 1, 1) = "," Then
                    If IsNumeric(Mid(strTempColor, 2, 2)) Then
                        ColorCode2 = Mid(strTempColor, 2, 2)
                        If Right(ColorCode2, 1) = "-" Or Right(ColorCode2, 1) = "+" Then
                            ColorCode2 = Mid(ColorCode2, 1, Len(ColorCode2) - 1)
                        End If
                        iColor = Mid(strTempColor, 1, 3)
                    Else
                        If IsNumeric(Mid(strTempColor, 2, 1)) Then
                            ColorCode2 = Mid(strTempColor, 2, 1)
                            iColor = Mid(strTempColor, 1, 2)
                        End If
                        
                    End If
                End If
            End If
            If ColorCode <> "" Then
                If COLOR2 = True Then
                    strColor(X) = Mid(strColor(X), Len(ColorCode) + Len(iColor) + 1)
                Else
                    strColor(X) = Mid(strColor(X), Len(ColorCode) + 1)
                End If
                
                i = InStr(ColorCode, ",")
                If i Then
                    'okay cabral irc doesn't use the background color
                    'let's take it out of the color code
                    ColorCode = Mid(ColorCode, 1, i - 1)
                End If
                'lets get the high color codes
                'down to a number between 0 and 15
                Do Until ColorCode <= 16
                    ColorCode = ColorCode - 16
                Loop
                RTF.SelColor = color(ColorCode)
                If ColorCode2 <> "" Then
                    Do Until ColorCode2 <= 16
                        ColorCode2 = ColorCode2 - 16
                    Loop
                    charf.crBackColor = color(ColorCode2)
                    ret = SendMessageLong(RTF.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(charf))
                End If
                'underline code = 
                'bold code = 
                If InStr(strColor(X), "") Or InStr(strColor(X), "") Then
                    strBold = Split(strColor(X), "")
                    RTF.SelBold = True
                    For i = 0 To UBound(strBold)
                        RTF.SelBold = Not (RTF.SelBold)
                        z = InStr(strBold(i), "")
                        If z Then
                            strUnderline = Split(strBold(i), "")
                            RTF.SelUnderline = True
                            For z = 0 To UBound(strUnderline)
                                RTF.SelUnderline = Not (RTF.SelUnderline)
                                RTF.SelText = strUnderline(z)
                            Next z
                        Else
                            RTF.SelText = strBold(i)
                        End If
                    Next i
                    'RTF.SelText = strColor(X)
                Else
                    RTF.SelText = strColor(X)
                End If
            Else
                'there are no color codes
                RTF.SelColor = color(colorX.normal)
                charf.crBackColor = RTF.BackColor
                ret = SendMessageLong(RTF.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(charf))
                If InStr(strColor(X), "") Or InStr(strColor(X), "") Then
                    strBold = Split(strColor(X), "")
                    RTF.SelBold = True
                    For i = 0 To UBound(strBold)
                        RTF.SelBold = Not (RTF.SelBold)
                        z = InStr(strBold(i), "")
                        If z Then
                            RTF.SelUnderline = True
                            strUnderline = Split(strBold(i), "")
                            For z = 0 To UBound(strUnderline)
                                RTF.SelUnderline = Not (RTF.SelUnderline)
                                RTF.SelText = strUnderline(z)
                            Next z
                        Else
                            RTF.SelText = strBold(i)
                        End If
                    Next i
                    'RTF.SelText = strColor(X)
                Else
                    RTF.SelText = strColor(X)
                End If
                'RTF.SelText = strColor(X)
            End If
        Next X
    Else
        'there are no color codes in the text
        'let's send it to the channel right away
        'check for bold first
        If InStr(strData, "") Then
            strBold = Split(strData, "")
            RTF.SelBold = True
            For i = 0 To UBound(strBold)
                RTF.SelBold = Not (RTF.SelBold)
                RTF.SelText = strBold(i)
            Next i
            'RTF.SelText = strColor(X)
        Else
            RTF.SelText = strData
        End If
        'RTF.SelText = strData
    End If

    RTF.SelColor = color(colorX.normal)
    charf.crBackColor = RTF.BackColor
    ret = SendMessageLong(RTF.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(charf))
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelText = vbCrLf
    
    LockWindowUpdate 0
    Exit Sub
PassColorCode:
    RTF.SelColor = color(colorX.normal)
    charf.crBackColor = RTF.BackColor
    ret = SendMessageLong(RTF.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(charf))
    RTF.SelBold = False
    RTF.SelUnderline = False
    RTF.SelText = strData
    RTF.SelText = vbCrLf
    
    LockWindowUpdate 0
    Exit Sub
End Sub




