Attribute VB_Name = "ColorRTF"
Option Explicit



Public Sub DoColor(RTF As vbalRichEdit, strColor As String)
    Dim i As Integer
    Dim a As Integer
    Dim b As Integer
    Dim b1 As Integer
    Dim b2 As Integer
    Dim ColorCode As Integer
    Dim BGColorCode As String
    Dim ColorFound As Boolean
    Dim BoldFound As Boolean
    
    Dim color(0 To 15) As Long
    color(0) = vbWhite 'white
    color(1) = vbBlack 'black
    color(2) = RGB(0, 0, 140) 'dark blue
    color(3) = RGB(0, 140, 0) 'dark green
    color(4) = vbRed 'red
    color(5) = RGB(110, 65, 0) 'brown
    color(6) = RGB(140, 0, 140) 'purple
    color(7) = RGB(248, 146, 0) 'orange
    color(8) = RGB(255, 255, 0) 'yellow
    color(9) = vbGreen 'light green
    color(10) = RGB(0, 140, 140) 'dark blue green
    color(11) = RGB(0, 255, 255) 'light blue green
    color(12) = vbBlue 'light blue
    color(13) = vbMagenta 'magenta
    color(14) = RGB(140, 140, 140) 'grey
    color(15) = RGB(200, 200, 200) 'light grey
    
    ColorFound = False
    BoldFound = False
    If InStr(strColor, "") < InStr(strColor, "") Then
        a = InStr(strColor, "")
        If a = 0 Then a = 1
        'MsgBox "COLOR"
    Else
        a = InStr(strColor, "")
        If a = 0 Then a = 1
        'MsgBox "BOLD"
    End If
    Do Until InStr(strColor, "") = 0 And InStr(strColor, "") = 0
    'msgBox strColor
    'For i = 1 To Len(strColor)
        'find the start of color code or bold code
        RTF.InsertContents SF_TEXT, Mid(strColor, 1, a - 1)
        b1 = InStr(strColor, "") 'color
        b2 = InStr(strColor, "")  'bold
        If b1 = 0 And b2 <> 0 Then
            b1 = b2 + 1
        Else
            If b2 = 0 And b1 <> 0 Then
                b2 = b1 + 1
            End If
        End If
        If b1 < b2 Then
            a = b1
            If a = 0 Then a = 1
            'MsgBox "COLOR"
        Else
            a = b2
            If a = 0 Then a = 1
            'MsgBox "BOLD"
        End If
        If strColor = "" Then Exit Do
        If Mid(strColor, a, 1) = "" Then
            'if it's another color Kode then just delete it and move on
            If Mid(strColor, a + 1, 1) = "" Then
                strColor = "" & Mid(strColor, a + 2)
                i = i - 1
            End If
            'get first number in color code
            If IsNumeric(Mid(strColor, a + 1, 1)) Then
                ColorCode = Mid(strColor, a + 1, 1)
                ColorFound = True
                If IsNumeric(Mid(strColor, a + 2, 1)) Then
                    ColorCode = ColorCode & Mid(strColor, a + 2, 1)
                    strColor = Mid(strColor, a + 3)
                    'loop until colorcode less than 15
                    Do Until ColorCode < 16
                        If ColorCode > 15 Then
                            ColorCode = ColorCode - 15
                        End If
                    Loop
                    Debug.Print "2: " & Mid(strColor, 1, 1)
                    'START of background color if colorcode LEN = 2
                    If Mid(strColor, 1, 1) = "," Then
                        If IsNumeric(Mid(strColor, 2, 1)) Then
                            'first color code # on bg color
                            BGColorCode = BGColorCode & Mid(strColor, 2, 1)
                            strColor = Mid(strColor, 3)
                            If IsNumeric(Mid(strColor, 1, 1)) Then
                                BGColorCode = BGColorCode & Mid(strColor, 1, 1)
                                strColor = Mid(strColor, 2)
                            Else
                                'BG Color code is only one number
                            End If
                        Else
                            'there is no bg color
                        End If
                    End If
                    'loop until BG colorcode less than 15
                    If BGColorCode <> "" Then
                        Do Until BGColorCode < 16
                            If BGColorCode > 15 Then
                                BGColorCode = BGColorCode - 15
                            End If
                        Loop
                    End If
                    If BGColorCode <> "" Then
                        RTF.FontBackColour = color(BGColorCode)
                    End If
                    'reset Back Ground Color
                    BGColorCode = ""
                    'END BG COLOR
                    RTF.FontColour = color(ColorCode)
                    ColorFound = False
                Else
                    'if not a second digit
                    RTF.FontColour = color(ColorCode)
                    strColor = Mid(strColor, a + 2)
                    Debug.Print "1: " & Mid(strColor, 1, 1)
                    'start of background color if colorcode LEN = 1
                    If Mid(strColor, 1, 1) = "," Then
                        If IsNumeric(Mid(strColor, 2, 1)) Then
                            'first color code # on bg color
                            BGColorCode = BGColorCode & Mid(strColor, 2, 1)
                            strColor = Mid(strColor, 3)
                            If IsNumeric(Mid(strColor, 1, 1)) Then
                                BGColorCode = BGColorCode & Mid(strColor, 1, 1)
                                strColor = Mid(strColor, 2)
                            Else
                                'BG Color code is only one number
                            End If
                        Else
                            'there is no bg color
                        End If
                    End If
                    'loop until BG colorcode less than 15
                    If BGColorCode <> "" Then
                        Do Until BGColorCode < 16
                            If BGColorCode > 15 Then
                                BGColorCode = BGColorCode - 15
                            End If
                        Loop
                    End If
                    If BGColorCode <> "" Then
                        RTF.FontBackColour = color(BGColorCode)
                    End If
                    'reset Back Ground Color
                    BGColorCode = ""
                End If
            Else
                'No Color code after initial color code
                RTF.FontColour = color(1)
                strColor = Mid(strColor, a + 1)
            End If
        Else
            If Mid(strColor, a, 1) = "" Then
                RTF.FontBold = BoldFound
                strColor = Mid(strColor, a + 1)
                BoldFound = Not (BoldFound)
                'MsgBox BoldFound
                b1 = InStr(strColor, "") 'color
                b2 = InStr(strColor, "")  'bold
                If b1 = 0 And b2 <> 0 Then
                    b1 = b2 + 1
                Else
                    If b2 = 0 And b1 <> 0 Then
                        b2 = b1 + 1
                    End If
                End If
                If b1 <> 0 Or b2 <> 0 Then
                    If b1 < b2 Then
                        RTF.InsertContents SF_TEXT, Mid(strColor, 1, b1 - 1)
                    Else
                        RTF.InsertContents SF_TEXT, Mid(strColor, 1, b2 - 1)
                    End If
                Else
                    RTF.InsertContents SF_TEXT, Mid(strColor, 1)
                End If
            End If
            'strColor = Mid(strColor, a + 1)
            'i = 0
            'DoEvents
        End If
    'Next i
    Loop
    RTF.InsertContents SF_TEXT, strColor
    RTF.FontBold = False
    RTF.FontColour = vbBlack
    RTF.InsertContents SF_TEXT, vbCrLf

End Sub

