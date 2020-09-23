VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Echo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Echo(strTarget As String, strData As String)
    Dim i As Integer
    If LCase(strTarget) = "status" Then
        RTF = frmStatus.txtStatus
    End If
    For i = 1 To 10
        If LCase(RTF) = LCase(chanstats(i).name) Then
            RTF = channel(i).txtText
        End If
    Next i
    On Error GoTo PassColorCode
    
    'lock rtf
    LockWindowUpdate RTF.hWnd

    'chr(2) = bold
    'chr(3) = color
    'Chr(31) = Underline

    'haven't figured what this one does, it's pretty useless
    'so we'll delete this irc text code so it doesn't show up
    strData = Replace(strData, "", "")
    
    Dim color(0 To 15) As Long
    color(0) = vbWhite 'white
    color(1) = vbBlack 'black
    color(2) = RGB(0, 0, 140) 'dark blue
    color(3) = RGB(0, 140, 0) 'dark green
    color(4) = vbRed 'red
    color(5) = RGB(110, 65, 0) 'brown
    color(6) = RGB(140, 0, 140) 'purple
    color(7) = RGB(248, 146, 0) 'orange
    color(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    color(9) = vbGreen 'light green
    color(10) = RGB(0, 140, 140) 'dark blue green
    color(11) = RGB(0, 255, 255) 'light blue green
    color(12) = vbBlue 'light blue
    color(13) = vbMagenta 'magenta
    color(14) = RGB(140, 140, 140) 'grey
    color(15) = RGB(200, 200, 200) 'light grey
    
    Dim strColor() As String
    Dim strBold() As String
    Dim strUnderline() As String
    Dim strTempColor As String
    Dim ColorCode As String
    Dim ColorCode2 As String
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
                Do Until ColorCode < 16
                    ColorCode = ColorCode - 15
                Loop
                RTF.SelColor = color(ColorCode)
                If ColorCode2 <> "" Then
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

