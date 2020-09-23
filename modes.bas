Attribute VB_Name = "modes"
Option Explicit
Global MyModes As String

Public Sub OP(strValue As String, username As String, target As String, strModeName As Variant)
    'MsgBox "OP: " & strValue & " " & username & " wants to OP " & strModeName & " in " & target
    Dim i As Integer
    Dim X As Integer
    username = Replace(username, ":", "")
    For i = 1 To ChannelMax
        If LCase(ChannelName(i)) = LCase(target) Then
            For X = 0 To channel(i).lstNames.ListCount - 1
                If strValue = "+" Then
                    If LCase(channel(i).lstNames.List(X)) = LCase(strModeName) Then
                        channel(i).lstNames.RemoveItem (X)
                        channel(i).lstNames.AddItem "@" & strModeName
                        If iShow.modes Then
                            Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " ops " & strModeName)
                        Else
                            Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " ops " & strModeName & " in " & target & vbCrLf & "-")
                        End If
                    End If
                Else 'if value = "-"
                    If LCase(channel(i).lstNames.List(X)) = "@" & LCase(strModeName) Then
                        channel(i).lstNames.RemoveItem (X)
                        channel(i).lstNames.AddItem strModeName
                        If iShow.modes Then
                            Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " deops " & strModeName)
                        Else
                            Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " deops " & strModeName & " in " & target & vbCrLf & "-")
                        End If
                    End If
                End If
            Next X
        End If
    Next i

End Sub
Public Sub VOICE(strValue As String, username As String, target As String, strModeName As Variant)
    'MsgBox "OP: " & strValue & " " & username & " wants to VOICE " & strModeName & " in " & target
    Dim i As Integer
    Dim X As Integer

    For i = 1 To ChannelMax
        If LCase(ChannelName(i)) = LCase(target) Then
            For X = 0 To channel(i).lstNames.ListCount - 1
                If strValue = "+" Then
                    If LCase(channel(i).lstNames.List(X)) = LCase(strModeName) Then
                        channel(i).lstNames.RemoveItem (X)
                        channel(i).lstNames.AddItem "+" & strModeName
                        If iShow.modes Then
                            Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " adds a voice to " & strModeName)
                        Else
                            Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " adds a voice to " & strModeName & " in " & target & vbCrLf & "-")
                        End If
                    End If
                Else 'if value = "-"
                    If LCase(channel(i).lstNames.List(X)) = "+" & LCase(strModeName) Then
                        channel(i).lstNames.RemoveItem (X)
                        channel(i).lstNames.AddItem strModeName
                        If iShow.modes Then
                            Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " devoiced " & strModeName)
                        Else
                            Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " devoiced " & strModeName & " in " & target & vbCrLf & "-")
                        End If
                    End If
                End If
            Next X
        End If
    Next i
End Sub
Public Sub INVISIBLE(strValue As String, username As String, target As String)
    'MsgBox "OP: " & strValue & " " & username & " wants to HIDE " & " in " & target
    If strValue = "+" Then
        MyModes = MyModes & "i"
    Else
        MyModes = Replace(MyModes, "i", "")
    End If
    
    frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
    Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & nickname & " sets mode " & strValue & "i" & vbCrLf & "-")
End Sub
Public Sub BAN(strValue As String, username As String, target As String, strModeName As Variant)
    'MsgBox "OP: " & strValue & " " & username & " wants to HIDE " & strModeName & " in " & target
    Dim i As Integer
     For i = 1 To ChannelMax
        If LCase(ChannelName(i)) = LCase(target) Then
            If strValue = "+" Then
                If iShow.modes Then
                    Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " bans " & strModeName)
                Else
                    Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " bans " & strModeName & " in " & target & vbCrLf & "-")
                End If
            Else 'if value = "-"
                If iShow.modes Then
                    Call DoColor(channel(i).txtText, "" & color.mode & "*** " & username & " unbans " & strModeName)
                Else
                    Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & username & " unbans " & strModeName & " in " & target & vbCrLf & "-")
                End If
            End If
        End If
    Next i
End Sub
Public Sub LIMIT(strValue As String, username As String, target As String, strModeName As Variant)
    Dim X As Integer
    Dim i As Integer

    Select Case strValue
        Case "+"
            For i = 1 To ChannelMax
                If LCase(ChannelName(i)) = LCase(target) Then
                        ChannelLimit(i) = strModeName
                        ChannelModes(i) = Replace(ChannelModes(i), "l", "")
                        ChannelModes(i) = ChannelModes(i) & "l"
                        'channel(i).Caption = target & " [+" & ChannelModes(i) & " " & ChannelLimit(i) & "] : " & ChannelTopic(i)
                        UpdateCaption i
                    Exit For
                End If
            Next i
        Case "-"
            MsgBox "Minus: " & username & " wants to set a limit of " & strModeName & " in channel " & target
    End Select
End Sub
Public Sub ChangeTopic(xUsername As String, xChannel As String, xTopic As String)
    Dim i As Integer
     For i = 1 To ChannelMax
        If LCase(ChannelName(i)) = LCase(xChannel) Then
            channel(i).txtTopic = xTopic
            'script
            chanstats(i).topic = xTopic
            Call DoColor(channel(i).txtText, "" & color.topic & "*** " & xUsername & " changes topic to '" & xTopic & "'")
        End If
    Next i
End Sub
Sub REGISTER(strValue As String, username As String, target As String)
    'MsgBox "OP: " & strValue & " " & username & " wants to HIDE " & " in " & target
    If strValue = "+" Then
        MyModes = MyModes & "r"
    Else
        MyModes = Replace(MyModes, "r", "")
    End If
    
    frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server
    Call DoColor(frmStatus.txtStatus, "" & color.mode & "*** " & nickname & " sets mode " & strValue & "r" & vbCrLf & "-")
End Sub
