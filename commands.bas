Attribute VB_Name = "commands"
Option Explicit
Public Const WM_PASTE = &H302

Public Sub command(username As String, command As String, target As String, parms As String)
    
    Dim timestamp As String
    timestamp = Now



    Dim colorX(0 To 15) As Long
    colorX(0) = vbWhite 'white
    colorX(1) = vbBlack 'black
    colorX(2) = RGB(0, 0, 140) 'dark blue
    colorX(3) = RGB(0, 140, 0) 'dark green
    colorX(4) = vbRed 'red
    colorX(5) = RGB(110, 65, 0) 'brown
    colorX(6) = RGB(140, 0, 140) 'purple
    colorX(7) = RGB(248, 146, 0) 'orange
    colorX(8) = vbYellow 'RGB(200, 200, 100)   'yellow
    colorX(9) = vbGreen 'light green
    colorX(10) = RGB(0, 140, 140) 'dark blue green
    colorX(11) = RGB(0, 255, 255) 'light blue green
    colorX(12) = vbBlue 'light blue
    colorX(13) = vbMagenta 'magenta
    colorX(14) = RGB(140, 140, 140) 'grey
    colorX(15) = RGB(200, 200, 200) 'light grey




    Dim i As Integer
    Dim X As Integer
    Dim found As Boolean
    
    
    'chop off the ":" in front of all parameters
    parms = LTrim(parms)
    If left(parms, 1) = ":" Then
        parms = Mid(parms, 2)
    End If
    Debug.Print parms
    If left(target, 1) = ":" Then
        target = Mid(target, 2)
    End If

    'uncomment to show on status
    'frmStatus.txtStatus.SelText = command & " " & parms & vbCrLf & "TARGET: " & target & vbCrLf
    'MsgBox command & " " & parms & vbCrLf & "TARGET: " & target & vbCrLf
    
    'break down usrename
    Dim UserEmail As String
    'this is for if the command wasn't sent by a user
    'it's most likely a server command
    Dim blServer As Boolean
    blServer = True
    For i = 1 To Len(username)
        If Mid(username, i, 1) = "!" Then
            UserEmail = Mid(username, i + 1)
            'the user has an email so it's not a server message
            blServer = False
            username = Mid(username, 1, i - 1)
            If left(username, 1) = ":" Then
                username = Mid(username, 2)
                'DoColor frmStatus.txtStatus, "5USER EMAIL: " & UserEmail
                'DoColor frmStatus.txtStatus, "5USER NAME: " & username
            End If
        End If
    Next i
    
    'IF URL CATCHER is ON then we'll check to see if there is a web address in there
    If InStr(LCase(parms), "http://") Then
        'URL found
        Dim strURL() As String
        Dim strURL2() As String
        strURL = Split(LCase(parms), " ")
        For i = 0 To UBound(strURL)
            If LCase(left(strURL(i), 4)) = "http" Then
                strURL2 = Split(strURL(i), Chr(32))
                    addURL strURL(i), username, timestamp
            End If
        Next i
    End If
    
    'this is for custom text
    Dim strCText As String
    
    Select Case UCase(command)
        Case "JOIN"
            For i = 1 To ChannelMax
                'if you join channel then load channel window
                If ChannelName(i) = "" Then
                    If LCase(username) = LCase(nickname) Then
                        Load channel(i)
                        channel(i).SHOW
                        'set colors that user wants
                        channel(i).txtSend.BackColor = colorX(color.bgText)
                        channel(i).txtText.BackColor = colorX(color.bgText)
                        channel(i).txtSend.ForeColor = colorX(color.normal)
                        channel(i).lstNames.ForeColor = colorX(color.normal)
                        channel(i).lstNames.BackColor = colorX(color.bgText)
                        '
                        ChannelName(i) = LCase(target)
                        ChannelModes(i) = ""
                        channel(i).Caption = target
                        'add users
                         UpdateCaption i
                        
                        mdiMain.tcp.SendData "MODE " & target & vbCrLf
                        
                        DoColor channel(i).txtText, "" & color.join & "*** Now talking in " & target
                        'add to status bar
                        Call AddTaskbar(target, 2)
                        
                        
                        ''''''add it to treeview
                        mdiMain.tvMain.Nodes.Add "mainChannels", tvwChild, LCase(target), target, 4
                        mdiMain.tvMain.Nodes.Item(3).Expanded = True
                        Exit For
                    End If
                End If
                'other users join channel
                'update nick list box
                If LCase(ChannelName(i)) = LCase(target) Then
                    channel(i).lstNames.AddItem username
                    'count users
                    UpdateCaption i
                    'custom text
                    strCText = strCustom.join
                    strCText = Replace(strCText, "$chan", target)
                    strCText = Replace(strCText, "$nick", username)
                    strCText = Replace(strCText, "$address", UserEmail)
                    strCText = Replace(strCText, "$color", "")
                    strCText = Replace(strCText, "$time", timestamp)
                    If iShow.joinpart Then
                        'show in channel that user joined channel
                        Call DoColor(channel(i).txtText, "" & color.join & strCText)
                    Else
                        'show in status window
                        Call DoColor(frmStatus.txtStatus, "" & color.join & strCText)
                        frmStatus.txtStatus.SelText = "-" & vbCrLf
                        Exit For
                    End If
                End If
            Next i
        Case "PART"
            events.chanpart = target
            events.nickpart = username
            Debug.Print target & " - " & username
            For i = 1 To ChannelMax
                'if you leave then unload channel window
                If ChannelName(i) = LCase(target) Then
                    If LCase(username) = LCase(nickname) Then
                        Unload channel(i)
                        'clear channel name array
                        ChannelName(i) = ""
                        'delete from treeview
                        RemoveNode target
                        Exit For
                    End If
                End If
                'other users leave channel
                'update nick list box
                'If LCase(ChannelName(i)) = LCase(target) And username <> "" Then
                If LCase(ChannelName(i)) = LCase(target) Then
                    'cycle through names and remove if found
                    For X = 0 To channel(i).lstNames.ListCount - 1
                        If channel(i).lstNames.ListIndex Then
                            If LCase(channel(i).lstNames.List(X)) = LCase(username) Or LCase(channel(i).lstNames.List(X)) = LCase("@" & username) Or LCase(channel(i).lstNames.List(X)) = LCase("+" & username) Then
                                channel(i).lstNames.RemoveItem (X)
                                'add users
                                UpdateCaption i
                                'custom text
                                strCText = strCustom.part
                                strCText = Replace(strCText, "$chan", target)
                                strCText = Replace(strCText, "$nick", username)
                                strCText = Replace(strCText, "$address", UserEmail)
                                strCText = Replace(strCText, "$color", "")
                                strCText = Replace(strCText, "$time", timestamp)
                                If iShow.joinpart Then
                                    'show in channel that user left channel
                                    Call DoColor(channel(i).txtText, "" & color.part & strCText)
                                Else
                                    'show in status
                                    Call DoColor(frmStatus.txtStatus, "" & color.part & strCText)
                                    frmStatus.txtStatus.SelText = "-" & vbCrLf
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                    
                End If
            Next i
        Case "PRIVMSG"
            If left(target, 1) = "#" Then
                For i = 1 To ChannelMax
                    If ChannelName(i) = LCase(target) Then
                        'if it's a channel action
                        If LCase(left(parms, 7)) = LCase("ACTION") Then
                            'take off  at the end
                            If Right(parms, 1) = "" Then
                                parms = Mid(parms, 1, Len(parms) - 1)
                            End If
                            'Channel(i).txtText.SelText = " " & Mid(parms, 8) & vbCrLf
                            Call DoColor(channel(i).txtText, "" & color.action & "* " & username & " " & Mid(parms, 8))
                        Else
                            'DoColor with parms
                            'custom text
                            strCText = strCustom.pm
                            If Trim(strCText) <> "" Then
                                strCText = Replace(strCText, "$chan", target)
                                strCText = Replace(strCText, "$nick", username)
                                strCText = Replace(strCText, "$address", UserEmail)
                                strCText = Replace(strCText, "$color", "")
                                strCText = Replace(strCText, "$msg", parms)
                                strCText = Replace(strCText, "$time", timestamp)
                                Call DoColor(channel(i).txtText, "" & color.normal & strCText)
                            Else
                                Call DoColor(channel(i).txtText, "" & color.normal & "<" & color.whois & "" & username & "" & color.normal & "> " & parms)
                            End If
                            
                        End If
                    Else
                        'if exited channel before PART msg then show channel text in status
                        'bottom line echos 10 times in status (fix)
                        'frmStatus.txtStatus.SelText = target & ": <" & username & "> " & parms & vbCrLf
                        'Exit For
                    End If
                Next i
            Else
                'if someone wants to find out what you're using
                If LCase(left(parms, 8)) = LCase("VERSION") Then
                    Call ctcp("VERSION SENT", frmStatus.txtStatus, username)
                    'exit sub so you won't get an extra query message
                    GoTo DONE
                End If
                        
                'found = if user already has a query window open
                found = False

                For i = 1 To 100
                    If LCase(username) = LCase(QueryName(i)) Then
                        found = True
                        Exit For
                    End If
                Next i
                If found = False Then
                    For i = 1 To 100
                        If QueryName(i) = "" Then
                            'load query window
                            Load Query(i)
                            Query(i).SHOW
                            'set user colors
                            Query(i).txtSend.BackColor = colorX(color.bgText)
                            Query(i).txtSend.ForeColor = colorX(color.normal)
                            Query(i).txtQuery.BackColor = colorX(color.bgText)
                            '
                            Query(i).Caption = username & " [" & UserEmail & "]"
                            QueryName(i) = username
                            Call AddTaskbar(username, 1)
                            UserOn Query(i).txtQuery, username
                            
                            
                            '''add to treeview
                            mdiMain.tvMain.Nodes.Add "mainChats", tvwChild, LCase(username), username, 3
                            mdiMain.tvMain.Nodes.Item(2).Expanded = True
                            Exit For
                        End If
                    Next i
                End If
                
                For i = 1 To 100
                    If LCase(QueryName(i)) = LCase(username) Then

                        If left(parms, 1) = "" Then
                            Call ctcp(parms, Query(i).txtQuery, username)
                            Query(i).Caption = username & " [" & UserEmail & "]"
                        Else
                            Call DoColor(Query(i).txtQuery, "" & color.normal & "<" & color.whois & "" & username & "" & color.normal & "> " & parms)
                            Query(i).Caption = username & " [" & UserEmail & "]"
                        End If
                    End If
                Next i
            End If
        Case "NICK"
            'Update your nickname variable
            If LCase(username) = LCase(nickname) Then
                nickname = target
                frmStatus.Caption = "Status: " & target & " on " & server
            End If

                'someone changed nickname
                'cycle through each window since the NICK message
                'tells you no channel or anything
                For i = 1 To ChannelMax
                    'if channel exsists (should have name)
                    If ChannelName(i) <> "" Then
                        'cycle through each listbox on each channel window
                        For X = 0 To channel(i).lstNames.ListCount - 1
                            'if NICK name matches lstnames(name) then remove
                            If LCase(channel(i).lstNames.List(X)) = LCase(username) Then
                                'remove from list box for certain window
                                channel(i).lstNames.RemoveItem (X)
                                'add new nickname
                                channel(i).lstNames.AddItem target
                                'Display nick change in channel
                                DoColor channel(i).txtText, "" & color.nick & "*** " & username & " has changed nickname to " & target
                                Else
                                    If LCase(channel(i).lstNames.List(X)) = LCase("@" & username) Then
                                        'remove from list box for certain window
                                        channel(i).lstNames.RemoveItem (X)
                                        'add new nickname
                                        channel(i).lstNames.AddItem "@" & target
                                        'Display nick change in channel
                                        DoColor channel(i).txtText, "" & color.nick & "*** " & username & " has changed nickname to " & target
                                    Else
                                        If LCase(channel(i).lstNames.List(X)) = LCase("+" & username) Then
                                            'remove from list box for certain window
                                            channel(i).lstNames.RemoveItem (X)
                                            'add new nickname
                                            channel(i).lstNames.AddItem "+" & target
                                            'Display nick change in channel
                                            DoColor channel(i).txtText, "" & color.nick & "*** " & username & " has changed nickname to " & target
                                        End If
                                    End If
                                End If
                            
                        Next X
                    End If
                    For X = 1 To 100
                        If LCase(QueryName(X)) = LCase(username) Then
                            QueryName(X) = target
                            Call RemoveTaskbar(username)
                            Call AddTaskbar(target, 1)
                            Query(X).Caption = target & " [" & UserEmail & "]"
                            '''
                            RemoveNode username
                            mdiMain.tvMain.Nodes.Add "mainChats", tvwChild, LCase(target), target
                        End If
                        Exit For
                    Next X
                Next i
        Case "NOTICE"
            If blServer Then
                Call DoColor(frmStatus.txtStatus, "" & color.notice & "NOTICE: " & parms)
            End If
            'ctcp
            If LCase(left(parms, 1)) = LCase("") Then
                Call ctcp(parms, frmStatus.txtStatus, username)
                'exit sub so you won't get an extra query message
                GoTo DONE
            End If

            'channel notice
            If left(target, 1) = "#" Then
                For i = 1 To ChannelMax
                    'cycle through channel names and if found a match then
                    'writing to channel window
                    If LCase(target) = LCase(ChannelName(i)) Then
                        DoColor channel(i).txtText, "" & color.notice & target & ": <" & username & "> " & parms
                    Else
                        'show channel NOTICE to status window
                        
                    End If
                Next i
            Else
                'if it was not a server notice then, continue for a personal notice
                If blServer = False Then
                    'personal NOTICE to you
                    'query window for username not found yet
                    found = False
                    If left(LCase(mdiMain.ActiveForm.Caption), 7) = LCase("Status:") Then
                        Call DoColor(frmStatus.txtStatus, "" & color.notice & username & ": " & parms & vbCrLf & "-")
                    Else
                        If left(mdiMain.ActiveForm.Caption, 1) = "#" Then
                            'in a channel
                            Call DoColor(mdiMain.ActiveForm.txtText, "" & color.notice & username & ": " & parms)
                        Else
                            'display back to status
                            Call DoColor(frmStatus.txtStatus, "" & color.notice & username & ": " & parms & vbCrLf & "-")
                        End If
                    End If
                End If
            End If
        Case "QUIT"
            'cycle through each window since the QUIT message
            'tells you no channel or anything
            For i = 1 To ChannelMax
                'if channel exsists (should have name)
                If ChannelName(i) <> "" Then
                    'cycle through each listbox on each channel window
                    For X = 0 To channel(i).lstNames.ListCount - 1
                        'if QUIT name matches lstnames(name) then remove
                        If LCase(channel(i).lstNames.List(X)) = LCase(username) Or LCase(channel(i).lstNames.List(X)) = LCase("@" & username) Or LCase(channel(i).lstNames.List(X)) = LCase("+" & username) Then
                            'remove from list box for certain window
                            channel(i).lstNames.RemoveItem (X)
                            'add users
                            UpdateCaption i
                            'custom text
                            strCText = strCustom.quit
                            strCText = Replace(strCText, "$reason", parms)
                            strCText = Replace(strCText, "$nick", username)
                            strCText = Replace(strCText, "$address", UserEmail)
                            strCText = Replace(strCText, "$color", "")
                            strCText = Replace(strCText, "$time", timestamp)
                            If iShow.quits Then
                                'Display quit in channel
                                Call DoColor(channel(i).txtText, "" & color.quit & strCText)
                            Else
                                'show in status instead
                                Call DoColor(frmStatus.txtStatus, "" & color.quit & strCText)
                                frmStatus.txtStatus.SelText = "-" & vbCrLf
                                Exit For
                            End If
                        End If
                    Next X
                End If
            Next i
        Case "KICK"
            'MsgBox target & " = Channel"
            'MsgBox username & "= OP"
            Dim word(1) As String
            i = InStr(2, parms, Chr(32))
            'Debug.Print i & parms
            word(1) = Trim(Mid(parms, 1, i - 1))
            'Debug.Print word(1) & " = Kicked"
            parms = Mid(parms, 2 + i)
            For i = 1 To ChannelMax
                If LCase(target) = LCase(ChannelName(i)) Then
                    'custom text
                    strCText = strCustom.kick
                    strCText = Replace(strCText, "$reason", parms)
                    strCText = Replace(strCText, "$nick", username)
                    strCText = Replace(strCText, "$address", UserEmail)
                    strCText = Replace(strCText, "$chan", target)
                    strCText = Replace(strCText, "$kicked", word(1))
                    strCText = Replace(strCText, "$color", "")
                    strCText = Replace(strCText, "$time", timestamp)
                    If iShow.kicks Then
                        Call DoColor(channel(i).txtText, "" & color.kick & strCText)
                    Else
                        Call DoColor(frmStatus.txtStatus, "" & color.kick & strCText & vbCrLf & "-")
                    End If
                    For X = 0 To channel(i).lstNames.ListCount - 1
                        If LCase(channel(i).lstNames.List(X)) = LCase(word(1)) Or _
                        LCase(channel(i).lstNames.List(X)) = LCase("@" & word(1)) Or _
                        LCase(channel(i).lstNames.List(X)) = LCase("+" & word(1)) Then
                            channel(i).lstNames.RemoveItem (X)
                            'add users
                            UpdateCaption i
                        End If
                    Next X
                End If
            Next i
        Case "MODE"
            'MsgBox username & " has " & parms & " on " & target
        
            Dim strWord() As String
            strWord = Split(parms, " ")
            If Len(strWord(0)) > 2 Then
                'execute if more than one mode is being made at a time
                'for example +oo-vv name name name name
                'irc only allows four possible mode changes at once
                Dim strMode(1 To 4) As String
                Dim CurrentMode As String
                X = 1
                strWord(0) = Trim(strWord(0))
                For i = 1 To Len(strWord(0))
                    If strWord(0) = "" Then Exit For
                    If left(strWord(0), 1) = "+" Or left(strWord(0), 1) = "-" Then
                        strMode(X) = Mid(strWord(0), 1, 2)
                        CurrentMode = Mid(strWord(0), 1, 1)
                        strWord(0) = Mid(strWord(0), 3)
                    Else
                        strMode(X) = CurrentMode & Mid(strWord(0), 1, 1)
                        strWord(0) = Mid(strWord(0), 2)
                    End If
                    X = X + 1
                    i = 1
                Next i
            Else
                'This else is executed if the mode only contains
                'one mode like: +o
                'strWord(1) = the name that was affected by the mode
                Select Case LCase(Mid(strWord(0), 2))
                    Case "o"
                        Call OP(left(strWord(0), 1), username, target, strWord(1))
                    Case "v"
                        Call VOICE(left(strWord(0), 1), username, target, strWord(1))
                    Case "i"
                        Call INVISIBLE(left(strWord(0), 1), username, target)
                    Case "r"
                        Call REGISTER(left(strWord(0), 1), username, target)
                    Case "b"
                        Call BAN(left(strWord(0), 1), username, target, strWord(1))
                    Case "l"
                        Call LIMIT(left(strWord(0), 1), username, target, strWord(1))
                    Case "n"
                        MsgBox left(strWord(0), 1) & " " & username & " " & target & " " & strWord(1)
                    Case "t"
                        MsgBox left(strWord(0), 1) & " " & username & " " & target & " " & strWord(1)
                    Case Else
                        'DoColor frmStatus.txtStatus, "3*** MODE: " & username & ", " & command & ", " & target & ", " & parms
                        DoColor frmStatus.txtStatus, "" & color.mode & "*** " & Replace(username, ":", "") & " sets mode: " & parms & vbCrLf & "-"
                        'Left(strWord(0), 1) = + or - a mode
                        'mid(strword(0),2) = the mode letter
                        If left(strWord(0), 1) = "+" Then
                            MyModes = MyModes & Mid(strWord(0), 2)
                            frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
                        Else
                            MyModes = Replace(MyModes, Mid(strWord(0), 2), "")
                            frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
                        End If
                End Select
            End If
            'now we're gonna do the modes we got in the multiple modes part
            For i = 1 To 4
                If strMode(i) <> "" Then
                    Select Case LCase(Mid(strMode(i), 2))
                        Case "o"
                            Call OP(left(strMode(i), 1), username, target, strWord(i))
                        Case "v"
                            Call VOICE(left(strMode(i), 1), username, target, strWord(i))
                        Case "i"
                            Call INVISIBLE(left(strMode(i), 1), username, target)
                        Case "b"
                            Call BAN(left(strMode(i), 1), username, target, strWord(i))
                        Case "r"
                            Call REGISTER(left(strMode(i), 1), username, target)
                        Case Else
                            'DoColor frmStatus.txtStatus, "3*** MODE: " & username & ", " & command & ", " & target & ", " & parms
                            DoColor frmStatus.txtStatus, "" & color.mode & "*** " & Replace(username, ":", "") & " sets mode: " & parms & vbCrLf & "-"
                            'Left(strWord(0), 1) = + or - a mode
                            'mid(strword(0),2) = the mode letter
                            If left(strMode(i), 1) = "+" Then
                                MyModes = MyModes & Mid(strMode(i), 2)
                                frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
                            Else
                                MyModes = Replace(MyModes, Mid(strMode(i), 2), "")
                                frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
                            End If
                            'channelmodes
                            For X = 1 To ChannelMax
                                If LCase(ChannelName(i)) = LCase(target) Then
                                    If left(strMode(i), 1) = "-" Then
                                        
                                    Else
                                        'ChannelModes(i) = ChannelModes(i) = Replace(ChannelModes(i), Mid(strMode(i), 2), "")
                                        'ChannelModes(i) = ChannelModes(i) & Mid(strMode(i), 2)
                                        'channel(i).Caption = target & " [" & ChannelModes(i) & "]: " & ChannelTopic(i)
                                    End If
                                    Exit For
                                End If
                            Next X
                    End Select
                End If
            Next i
        Case "TOPIC"
            'channel topic is about to change
            'target = channel name to be changed in
            Call ChangeTopic(username, target, parms)
        Case "INVITE"
            DoColor frmStatus.txtStatus, "" & color.invite & "*** " & username & " invites you to join " & parms & vbCrLf & "-"
        '''''''''''''''''''''
        'These below are the odd ones, the one's that usually give you error messages
        'on why you can't connect to a server
        '''''''''''''''''''''
        

        Case "ERROR:"
            MsgBox parms
        Case "AUTH"
            DoColor frmStatus.txtStatus, "" & color.notice & "*** " & parms & vbCrLf & "-"
        Case ":CLOSING"
            DoColor frmStatus.txtStatus, "" & color.quit & "*** " & username & " " & command & " " & target & " " & parms & vbCrLf & "-"
        Case "DLINE"
            DoColor frmStatus.txtStatus, "" & color.quit & "*** " & username & " " & command & " " & target & " " & parms & vbCrLf & "-"
        Case Else
            DoColor frmStatus.txtStatus, "" & color.normal & "[04Username" & "" & color.normal & "]: " & username & vbCrLf & "" & color.normal & "[04Command" & "" & color.normal & "]: " & command & vbCrLf & "" & color.normal & "[04Target" & "" & color.normal & "]: " & target & vbCrLf & "" & color.normal & "[04Paramters" & color.normal & "]: " & parms & vbCrLf & "-" & vbCrLf
        End Select
DONE:
    Exit Sub
End Sub


Sub ctcp(parms As String, RTF As RichTextBox, username As String)
    Dim word() As String
    Dim X As Integer, Y As Integer, i As Integer
    parms = Mid(parms, 2)
    If Right(parms, 1) = "" Then
        DoColor frmStatus.txtStatus, "" & color.ctcp & "[" & username & " " & Mid(parms, 1, Len(parms) - 1) & "]" & vbCrLf & "-" & vbCrLf
    Else
        DoColor frmStatus.txtStatus, "" & color.ctcp & "[" & username & " " & parms & "]" & vbCrLf & "-" & vbCrLf
    End If
    'split the PARMS into seperate words
    'ReDim Preserve statement is the KEY
    word = Split(parms, Chr(32))
        'All done spliting each word up, word to the mothers
        Select Case UCase(word(0))
            Case "ACTION"
                parms = ""
                For X = 1 To UBound(word)
                    parms = parms & word(X) & Chr(32) 'chr(32 = blank space
                Next X
                Call DoColor(RTF, "" & color.action & "* " & username & " " & parms)
            Case "DCC"
                If UCase(word(1)) = UCase("CHAT") Then
                    If UCase(word(2)) = UCase("CHAT") Then
                        DoColor frmStatus.txtStatus, "" & color.ctcp & "* DCC Chat request from " & username '& " on his port " & word(5)
                        frmStatus.txtStatus.SelText = "-" & vbCrLf
                        'LIP = LongIP wod(4) in the dcc privmsg parm
                        Dim LIP As String
                        Dim wDCC As New frmDCCACCEPT
                        LIP = IrcGetIP(word(3))
                        'load dcc chat msg box?
                        word(4) = Replace(word(4), "", "")
                        wDCC.SHOW 0, mdiMain
                        wDCC.lblNickName = username
                        wDCC.lblIP = LIP
                        wDCC.lblPort = word(4)
                    End If
                End If
                If UCase(word(1)) = UCase("SEND") Then
                    '3 = filename
                    'Dim LIP As String
                    'LIP = IrcGetIP(word(4))
                    LIP = IrcGetIP(Val(word(UBound(word) - 2)))
                    
                    'setup new file transfer window
                    Dim NewFile As frmDCCFILE
                    Set NewFile = New frmDCCFILE
                                        
                    FileIndex = FileIndex + 1
                    If FileIndex > 999 Then FileIndex = 3
                    NewFile.Tag = FreeFile
                    'Load NewFile.FILE(FileIndex)
                    Load NewFile.FILE(NewFile.Tag)
 
                    
                    NewFile.SHOW 0, mdiMain
                    For i = 0 To UBound(word)
                        DoColor frmStatus.txtStatus, "5" & word(i)
                    Next i
                    
                    For i = 3 To UBound(word)
                        If IsNumeric(word(i)) = False Then
                            word(2) = word(2) & word(i) & "_"
                        Else
                            Exit For
                        End If
                    Next i

                    'word(2) = Mid(word(2), 1, Len(word(2)) - 1)
                    'word(2) = Replace(word(2), """", "")
                    NewFile.lblFile = word(2)
                    'frmDCCFILE.lblAddress = LIP & ":" & word(5)
                    NewFile.lblAddress = LIP & ":" & word(i + 1)
                    'frmDCCFILE.lblFileSize = word(6)
                    NewFile.lblFileSize = word(i + 2)
                    NewFile.lblNickName = username
                    NewFile.ProgressBar.Min = 0
                    NewFile.ProgressBar.Value = 0
                    NewFile.picComplete.BackColor = vbWhite
                    NewFile.ProgressBar.max = Val(word(UBound(word)))
                    NewFile.lblRCV = "Recieved: " & 0
                    NewFile.lblFilename = App.Path & "\downloads\" & username & "\" & frmDCCFILE.lblFile
                    
                    NewFile.Caption = "DCC File Get: " & word(2)
                End If
            Case "VERSION"
                'DoColor RTF, "4[" & username & " VERSION]" & vbCrLf & "-" & vbCrLf
                mdiMain.tcp.SendData "NOTICE " & username & " :VERSION Cabral IRC. Rockford Project. Jaime Cabral" & vbCrLf
            Case "PING"
                Dim PingTime As String
                For i = 1 To Len(word(1))
                    If IsNumeric(Mid(word(1), i, 1)) Then
                        PingTime = PingTime & Mid(word(1), i, 1)
                    Else
                        Exit For
                    End If
                Next i
                word(1) = PingTime
                If Right(word(1), 1) = Chr(1) Then word(1) = Mid(word(1), 1, Len(word(1)) - 1)
                X = Val(Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now)))) - Val(Int(word(1)))
                i = Int(X)
                'i = (X / 1000)
                'i = Int(i / 10)
                DoColor frmStatus.txtStatus, "" & color.ctcp & "[" & username & " PING Reply]: " & i & " seconds "
                'DoColor frmStatus.txtStatus, "" & color.ctcp & "Ping Reply: " & word(2) & " - " & Val(Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now))))
                If LCase(nickname) <> LCase(username) Then
                    mdiMain.tcp.SendData "NOTICE " & username & " :PING " & Val(Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now)))) & Chr(1) & vbCrLf
                    mdiMain.tcp.SendData "NOTICE " & username & " :I am running CircEX©" & vbCrLf
                Else
                    mdiMain.tcp.SendData "NOTICE " & username & " :I am running CircEX©" & vbCrLf
                End If
        End Select

        'exit sub so it won't go into the second part of this code (not implemented the 2nd part yet)

    
End Sub

