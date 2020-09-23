Attribute VB_Name = "numerics"
Option Explicit







Public Sub numeric(strServer As String, NUM As String, nickname As String, strLine As String)
    raw.NUM = NUM
    raw.server = strServer
    raw.nickname = nickname
    raw.parms = strLine

'On Error GoTo NumericError

    'uncomment to show on 'status
    'frmStatus.txtStatus.InsertContents SF_TEXT, num & " " & strLine & vbCrLf


    Dim i As Integer, X, ix As Integer
    Dim word() As String
    Dim strTemp As String
    Static ChannelCount As Integer
    Static ChannelPause As Integer
    Dim channame() As String

    server = strServer
    If MyModes <> "" Then
        frmStatus.Caption = "Status: [+" & MyModes & "] " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
    Else
        frmStatus.Caption = "Status: " & nickname & " on " & server & ":" & mdiMain.tcp.RemotePort
    End If
    'used for "322" to seperate info from PARMS
    Dim mItem As Variant
    Dim xChannelName As String
    Dim xUsers As String
    Dim xTopic As String
    'used for to open file
    Dim CNumber As Integer
    
    
    
    'if numeric requires spliting text then this will happen
    If NUM = "311" Or NUM = "312" Or NUM = "317" Then
        'split strLine into seperate words
        'ReDim Preserve statement is the KEY
        strTemp = strLine
        If InStr(strTemp, Chr(32)) Then
            Do Until InStr(strTemp, Chr(32)) = 0
                X = InStr(strTemp, Chr(32))
                If X Then
                    i = i + 1
                    ReDim Preserve word(i)
                    word(i) = Mid(strTemp, 1, X - 1)
                    strTemp = Mid(strTemp, X + 1)
                End If
            Loop
            ReDim Preserve word(i + 1)
            word(i + 1) = strTemp
        End If
    End If
    
    
    
    'IF URL CATCHER is ON then we'll check to see if there is a web address in there
    If InStr(LCase(strLine), "http://") Then
        'URL found
        Dim strURL() As String
        Dim strURL2() As String
        strURL = Split(LCase(strLine), " ")
        For i = 0 To UBound(strURL)
            If LCase(left(strURL(i), 4)) = "http" Then
                strURL2 = Split(strURL(i), Chr(32))
                    addURL strURL(i), Mid(strServer, 2), Now
            End If
        Next i
    End If
    
    
    Select Case NUM
        'MOTD 372 375 376 378
        Case "375"
            If iModes.i Then
                mdiMain.tcp.SendData "mode " & nickname & " +i" & vbCrLf
            End If
            If iModes.s Then
                mdiMain.tcp.SendData "mode " & nickname & " +s" & vbCrLf
            End If
            If iModes.w Then
                mdiMain.tcp.SendData "mode " & nickname & " +w" & vbCrLf
            End If
        Case "372"
            'frmMOTD.txtMOTD.SelText = strLine & vbCrLf
            If iShow.motd Then
                Call DoColor(frmMOTD.txtMOTD, "" & color.normal & ",00" & strLine)
                frmMOTD.SHOW
            Else
                Call DoColor(frmStatus.txtStatus, "" & color.normal & strLine)
            End If
            'DoEvents

            'DoEvents
        Case "376"
            'frmMOTD.SHOW
            'end MOTD
            Call DoColor(frmStatus.txtStatus, "" & color.normal & strLine & vbCrLf & "-" & vbCrLf)
        'NAMES 353
        Case "353"
            'channame(3) = first name in list
            'channame(1) = channel name
            'if names already is in the list then it won't add
            Dim FoundNAMES As Boolean
            'list /names #channel (if not in channel will return the list in status)
            Dim blSTATUS As Boolean
            blSTATUS = True
            FoundNAMES = False
            'seperate names
            channame = Split(strLine, " ")
            'Take off the ":" in front of first name
            channame(2) = Mid(channame(2), 2)
            'ChanName(1) = Channel Name
            'add name to channel listbox
            For i = 2 To UBound(channame)
                For ix = 1 To ChannelMax
                    If LCase(ChannelName(ix)) = LCase(channame(1)) Then
                        channame(i) = Replace(channame(i), "%", "")
                        channel(ix).lstNames.AddItem channame(i)
                        UpdateCaption ix
                    End If
                Next ix
            Next i
            DoColor frmStatus.txtStatus, "" & color.normal & strLine
        '366 /end NAMES
        Case "366"
            word = Split(strLine, " ")
            ACTION_CHANNEL = word(0)
            xINPUT "STATS", frmStatus.txtStatus
            Call DoColor(frmStatus.txtStatus, "" & color.normal & strLine & vbCrLf & "-" & vbCrLf)
        '321 Start List
        '322 list channels #channel #users :topic
        '323 end list
        Case "321"
            Call DoColor(frmStatus.txtStatus, "" & color.normal & "Listing Channels..." & vbCrLf & "-" & vbCrLf)
            ChannelCount = 0
            ChannelPause = 0
            'frmListChannels.SHOW
            frmChannels.SHOW
            frmChannels.lvwChan.ListItems.Clear
            Set frmChannels.lvwChan.SmallIcons = frmChannels.imgChan
            
            'use with old frmchannels
            'Unload frmChannels
            
            'lock list
            LockWindowUpdate frmChannels.lvwChan.hwnd
            
            'used with old frmchannels
            'Call SendMessage(frmChannels.lvwChan.hWnd, WM_SETREDRAW, 0, 0)
            

        Case "322"
        

            'add to channel list box
            'Tree View Control
            'frmChannels.ChannelView.Nodes.Add = strLine
            
            'seperate CHANNEL USERS :TOPIC
            'put in listview control (lvxChan control)
            word = Split(strLine, " ")
            xChannelName = word(0)
            xUsers = word(1)
            For i = 2 To UBound(word)
                xTopic = xTopic & word(i) & " "
            Next i
        If xChannelName <> "*" Then
            Set mItem = frmChannels.lvwChan.ListItems.Add(, , xChannelName)
            'Set mItem = frmChannels.lvwChan.ListItems.Add(, , xChannelName)
            mItem.SubItems(1) = xUsers
            mItem.SubItems(2) = xTopic
            
            ChannelCount = ChannelCount + 1
            ChannelPause = ChannelPause + 1
            If ChannelPause >= 125 Then
                ChannelPause = 0
                'frmChannels.Caption = "Cabral Channel List [" & ChannelCount & "]"
                'Call SendMessage(frmChannels.lvwChan.hWnd, WM_SETREDRAW, 1, 0)
                LockWindowUpdate 0
                frmListChannels.lstChannels.Refresh
                'Call SendMessage(frmChannels.lvwChan.hWnd, WM_SETREDRAW, 0, 0)
                LockWindowUpdate frmChannels.lvwChan.hwnd
            End If
        End If

        Case "323"
            'frmChannels.SHOW
            Call DoColor(frmStatus.txtStatus, "" & color.normal & "Finished listing channels! [" & ChannelCount & "]" & vbCrLf & "-" & vbCrLf)
            'unlock list
            'used with old frmchannels
            LockWindowUpdate 0
            frmListChannels.lstChannels.Refresh
            
            'used with old frmchannels
            frmChannels.Caption = "Cabral Channel List [" & ChannelCount & "]"
            'Call SendMessage(frmChannels.lvwChan.hWnd, WM_SETREDRAW, 1, 0)
            

        Case "332" 'TOPIC for channel shows when YOU JOIN
            'sample:  :www.ircserver.org 332 MyNickname #Channel :NO TOPIC NOW
            
            'lets split string
            Dim TopicArray() As String
            TopicArray = Split(strLine, " ")
            'channel name = TopicArray(0)
            'clear strline
            strLine = ""
            'rebuild strline/ the outcome will be the complete TOPIC and only topic
            For i = 1 To UBound(TopicArray)
                strLine = strLine & " " & TopicArray(i)
            Next i
            'remove the extra space you had when the above loop started
            strLine = LTrim(strLine)
            'remove leading ":"
            If left(strLine, 1) = ":" Then
                strLine = Mid(strLine, 2)
                Dim ChanTopic As String
                ChanTopic = strLine
            End If
            'search for channel to put topic in
            For X = 1 To ChannelMax
                If LCase(ChannelName(X)) = LCase(TopicArray(0)) Then
                    'we found the channel
                    'place in channel topic text box
                    channel(X).txtTopic = ""
                    ChannelTopic(X) = ""
                    Call DoColor(channel(X).txtTopic, "" & color.normal & strLine)
                    'topic textboxes tooltip
                    channel(X).txtTopic.ToolTipText = strLine
                    ChannelTopic(X) = channel(X).txtTopic.Text
                    ChannelTopic(X) = Replace(ChannelTopic(X), Chr(13), "")
                    ChannelTopic(X) = Replace(ChannelTopic(X), Chr(10), "")
                    channel(X).Caption = ChannelName(X) & " [+" & ChannelModes(X) & "] :" & ChannelTopic(X)
                    'script
                    chanstats(X).topic = strLine
                    'show topic in channel as you join
                    Call DoColor(channel(X).txtText, "" & color.join & "*** Topic is '" & ChanTopic & "'")
                    Exit For
                End If
            Next X
        Case "311"
            '311 RPL_WHOISUSER
            '"<nick> <user> <host> * :<real name>"
            Dim RealName As String
            For i = 4 To UBound(word)
                RealName = RealName & Chr(32) & word(i)
            Next i
            'frmStatus.txtStatus.TextRTF = frmStatus.txtStatus.TextRTF & word(1) & " is " & word(2) & "@" & word(3) & RealName & vbCrLf
            'mIRC style whois - must uncomment
            Call DoColor(frmStatus.txtStatus, "" & color.whois & word(1) & " is " & word(2) & "@" & word(3) & RealName)
            'My crappy ass style
            'frmStatus.txtStatus.SelText = "Nick Name: " & word(1) & vbCrLf
            'frmStatus.txtStatus.SelText = "Email: " & word(2) & "@" & word(3) & vbCrLf
            'frmStatus.txtStatus.SelText = "Real Name: " & RealName & vbCrLf
        Case "378"
            'frmMOTD.txtMOTD.SelText = strLine & vbCrLf
            'Beyond IRC likes to use this as a <nick> connecting from <real ip>
            Call DoColor(frmStatus.txtStatus, "" & color.whois & strLine)
            'Call DoColor(frmMOTD.txtMOTD, strLine)
        Case "312"
            '312 RPL_WHOISSERVER
            '"<nick> <server> :<server info>"
            Dim ServerQuote As String
            For i = 3 To UBound(word)
                ServerQuote = ServerQuote & Chr(32) & word(i)
            Next i
            'frmStatus.txtStatus.TextRTF = frmStatus.txtStatus.TextRTF & word(1) & " using " & word(2) & ServerQuote & vbCrLf
            Call DoColor(frmStatus.txtStatus, "" & color.whois & word(1) & " using " & word(2) & " [" & ServerQuote & "]")
            '313 RPL_WHOISOPERATOR
            '"<nick> :is an IRC operator"
        Case "317"
            '317 RPL_WHOISIDLE
            '"<nick> <integer> :seconds idle"
            Dim Seconds As Integer
            Dim Minutes As Variant
            Seconds = Val(word(2))
            Minutes = Seconds / 60
            For i = 1 To Len(Minutes)
                If Mid(Minutes, i, 1) = "." Then
                    Minutes = Mid(Minutes, 1, i - 1)
                End If
            Next i
            Seconds = Seconds - (Val(Minutes) * 60)
            If Val(Minutes) > 0 Then
                'mirc style
                Call DoColor(frmStatus.txtStatus, "" & color.whois & word(1) & " has been idle " & Minutes & " mins " & Seconds & " secs")
                'frmStatus.txtStatus.SelText = "Idle: " & Minutes & " mins " & Seconds & " secs" & vbCrLf
            Else
                'mirc style
                Call DoColor(frmStatus.txtStatus, "" & color.whois & word(1) & " has been idle " & Seconds & " secs")
                'frmStatus.txtStatus.SelText = "Idle: " & Seconds & " secs" & vbCrLf
            End If
        Case "318"
            '318 RPL_ENDOFWHOIS
            '"<nick> :End of /WHOIS list"
            Call DoColor(frmStatus.txtStatus, "" & color.whois & strLine & vbCrLf & "-")
        Case "307"
            'is a registered user
            Call DoColor(frmStatus.txtStatus, "" & color.whois & strLine)
        Case "319"
            'mirc style whois:  User is on these channels
            Call DoColor(frmStatus.txtStatus, "" & color.whois & Replace(strLine, ":", ""))
            '319 RPL_WHOISCHANNELS
            '"<nick> :{[@|+]<channel><space>}"
        Case "001"
            DoColor frmStatus.txtStatus, "" & color.server & strLine
            If iShow.channelfolder Then
                frmChannelFolder.SHOW 0, mdiMain
            End If
            If iShow.whoisnotify Then
                'mdiMain.tcp.SendData "WATCH C " & NOTIFYLIST & vbCrLf
                mdiMain.tcp.SendData "ISON " & RTrim(notifylist) & vbCrLf
            End If
            'now lets TRY to get your real ip
            word = Split(strLine)
            DoColor frmStatus.txtStatus, "*** Your IP is " & word(UBound(word))
        Case "002"
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "003"
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "004"
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "005"
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-" & vbCrLf
        'this is the server stats, sent to you on a /lusers
        '251 252 254 255 / 265 266
        Case "251"
            strLine = Replace(strLine, ":", "")
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "252"
             strLine = Replace(strLine, ":", "")
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "254"
            strLine = Replace(strLine, ":", "")
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "255"
            strLine = Replace(strLine, ":", "")
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-"
        Case "265"
            DoColor frmStatus.txtStatus, "" & color.server & strLine
        Case "266"
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-" & vbCrLf
        'error codes
        Case "401"
            'no such nick or channel
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-" & vbCrLf
        Case "433"
            'nickname is already in use
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-" & vbCrLf
        Case "482"
            'you are not channel operator
            DoColor frmStatus.txtStatus, "" & color.server & strLine & vbCrLf & "-" & vbCrLf
        Case "303"
            'NOTIFY IF USER IS ON
            'command sent to server is ISON <nick1> <nick2> <etc...>
            Dim NName() As String
            Dim xON As String
            Static tempNames As String
            strLine = Replace(strLine, ":", "")
            strLine = strLine & " "
            xON = tempNames
            If tempNames = strLine Then
                'same people are online, will not show that
                'DoColor frmStatus.txtStatus, "" & color.NOTIFY & "*** Same people are on IRC" & tempNames
            Else
                NName = Split(strLine, " ")
                For i = 0 To UBound(NName)
                    If NName(i) <> "" Then
                        X = InStr(tempNames & " ", NName(i) & " ")
                        If X Then
                            'DoColor frmStatus.txtStatus, "" & color.NOTIFY & "*** " & NName(i) & " is on IRC"
                            xON = Replace(xON, NName(i) & " ", "")
                            'DoColor frmStatus.txtStatus, "" & color.NOTIFY & "*** ON: " & xON & "!"
                        Else
                            'ok, so this is a new person that is on IRC - lets tell the user it is on
                            DoColor frmStatus.txtStatus, "" & color.notify & "*** " & NName(i) & " is on IRC"
                            If iShow.whoisnotify Then
                                mdiMain.tcp.SendData "WHOIS " & NName(i) & vbCrLf
                            End If
                        End If
                    End If
                Next i
                'NName = Split(xON, " ")
                'For i = 0 To UBound(NName)
                '    If Trim(NName(i)) <> "" Then
                '       DoColor frmStatus.txtStatus, "" & color.notify & "*** " & NName(i) & " has left IRC"
                '    End If
                'Next i
                
                ''''''''''''''''''''''''
                'old notify list
                ''''''''''''''''''''''''
                NName = Split(strLine, " ")
                frmFriends.lstFriends.Clear
                For i = 0 To UBound(NName)
                    frmFriends.lstFriends.AddItem NName(i)
                    frmFriends.Caption = UBound(NName) & " Friends Online"
                Next i
                ''''''''''''''''''''''''
                ''''''''''''''''''''''''
                'NName = Split(strLine, " ")
                'frmFriends.lstFriends.Clear
                'For i = 0 To UBound(NName)
                '    mdiMain.tvMain.Nodes.Add "mainFriends", tvwChild, LCase(NName(i)), NName(i)
                'Next i
            End If
            tempNames = strLine
        Case "328"
            '#channel :url
            word = Split(strLine, " ")
            word(1) = Mid(word(1), 2)
            For i = 1 To ChannelMax
                If LCase(ChannelName(i)) = LCase(word(0)) Then
                    DoColor channel(i).txtText, "" & color.join & "*** " & word(0) & " homepage is " & word(1)
                End If
            Next i
        Case "333"
            'who set the topic and when?
            Dim strTopicTime As String
            word = Split(strLine, " ")
                For i = 1 To ChannelMax
                If LCase(ChannelName(i)) = LCase(word(0)) Then
                    'MsgBox word(0) & CDate(word(2))
                    strTopicTime = irc_time(word(2))
                    DoColor channel(i).txtText, "" & color.join & "*** Topic was set by " & word(1) & " [" & strTopicTime & "]"
                End If
            Next i
        Case "364"
            '/links
        Case "324"
            word = Split(strLine, " ")
            For i = 1 To ChannelMax
                If LCase(ChannelName(i)) = LCase(word(0)) Then
                    ChannelModes(i) = Mid(word(1), 2)
                    'channel(i).Caption = ChannelName(i) & " [+" & ChannelModes(i) & "] :" & ChannelTopic(i)
                    UpdateCaption i
                    
                    'user limit
                    ChannelLimit(i) = word(2)
                    'channel(i).Caption = ChannelName(i) & " [+" & ChannelModes(i) & " " & ChannelLimit(i) & "] :" & ChannelTopic(i)
                    UpdateCaption i
                    Exit For
                End If
            Next i
        Case "329"
            word = Split(strLine, " ")
            word(1) = irc_time(word(1))
            For i = 1 To ChannelMax
                If LCase(ChannelName(i)) = LCase(word(0)) Then
                    DoColor channel(i).txtText, "" & color.join & "*** " & ChannelName(i) & " was created on " & word(1)
                    Exit For
                End If
            Next i
        Case Else
            DoColor frmStatus.txtStatus, "" & color.normal & "[04" & NUM & "" & color.normal & "]10" & strLine
    End Select
    
    'script
    'mdiMain.script.Run "OnNumeric", strServer, NUM, nickname, strLine
    
'NumericError:
'    If Err.Number <> 0 Then
'        frmStatus.txtStatus.InsertContents SF_TEXT, "["
'        frmStatus.txtStatus.FontBold = True
'        frmStatus.txtStatus.FontColour = RGB(0, 140, 0)
'        frmStatus.txtStatus.InsertContents SF_TEXT, "Error"
'        frmStatus.txtStatus.FontBold = False
'        frmStatus.txtStatus.FontColour = vbBlack
'        frmStatus.txtStatus.InsertContents SF_TEXT, "]" & vbCrLf
'        frmStatus.txtStatus.InsertContents SF_TEXT, "Server: " & server & vbCrLf & "Numeric: " & num & vbCrLf & "Nickname: " & NickName & vbCrLf & "String: " & strLine & vbCrLf
'        frmStatus.txtStatus.InsertContents SF_TEXT, "Description: " & Err.Description & " [" & Err.Number & "]" & vbCrLf & "-" & vbCrLf
'        'Resume Next
'    End If

End Sub

