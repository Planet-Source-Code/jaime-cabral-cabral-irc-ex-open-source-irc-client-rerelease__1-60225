Attribute VB_Name = "ColorRTF"
Option Explicit



Public Sub DoColor(RTF As vbalRichEdit, strColor As String)
    Dim i As Integer
    Dim i2 As Integer
    Dim x As Integer
    Dim x2 As Integer
    Dim ColorCode As String
    Dim ColorCode2 As String
    Dim strTemp As String
    
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
    

    If InStr(strColor, "") Then
        Do Until InStr(strColor, "") = 0
            i = InStr(strColor, "")
            If i > 0 Then
                If Len(Mid(strColor, 1, i - 1)) > 0 Then
                    If Val(ColorCode2) > 0 Then
                        x2 = Len(ColorCode) + Len(ColorCode2) + 1
                        RTF.InsertContents SF_TEXT, Mid(strColor, x2, i - x2)
                    Else
                        x2 = Len(ColorCode)
                        If x2 = 0 Then
                            RTF.InsertContents SF_TEXT, Mid(strColor, 1, i - 1)
                        Else
                            RTF.InsertContents SF_TEXT, Mid(strColor, x2 + 1, i - (x2 + 1))
                        End If
                    End If
                    strColor = Mid(strColor, i + 1)
                    ColorCode = ""
                    ColorCode2 = ""
                End If
                strTemp = Mid(strColor, 1, 5)
                For x = 1 To 3
                    If IsNumeric(Mid(strTemp, x, 1)) Then
                        ColorCode = ColorCode & Mid(strTemp, x, 1)
                        If ColorCode > 15 Then
                            Do Until ColorCode < 16
                                ColorCode = ColorCode - 15
                            Loop
                        End If
                        If Len(ColorCode) > 2 Then ColorCode = Right(ColorCode, 2)
                        strColor = Mid(strColor, x)
                        RTF.FontColour = color(ColorCode)
                    Else
                        If ColorCode = "" Then
                            ColorCode = color(1)
                            RTF.FontColour = ColorCode
                        End If
                        If Mid(strTemp, x, 1) = "," Then
                            strTemp = Mid(strTemp, 2 + Len(ColorCode))
                            x2 = x
                            For i2 = 1 To 3
                                If IsNumeric(Mid(strTemp, i2, 1)) Then
                                    ColorCode2 = ColorCode2 & Mid(strTemp, i2, 1)
                                    If Len(ColorCode2) > 2 Then ColorCode2 = Right(ColorCode2, 2)
                                    strColor = Mid(strColor, i2)
                                    RTF.FontBackColour = color(ColorCode2)
                                Else
                                    Exit For
                                End If
                            Next i2
                        End If
                        ColorCode = color(1)
                        Exit For
                    End If
                Next x
                
            End If
        Loop
    End If
    RTF.InsertContents SF_TEXT, strColor & vbCrLf
    RTF.FontColour = color(1)
    RTF.FontBackColour = color(0)
End Sub

