Attribute VB_Name = "general"
Option Explicit

Public Const ChannelMax = 20
Public channel(1 To ChannelMax) As New frmChannel
Public ChannelName(1 To ChannelMax) As String
Public ChannelTopic(1 To ChannelMax) As String
Public ChannelModes(1 To ChannelMax) As String
Public ChannelLimit(1 To ChannelMax) As String

Public Query(1 To 100) As New frmQuery
Public QueryName(1 To 100) As String

'This one is for when you send a /me in a channel
'it'll store the name of the channel to be use with "ME" case
'in xINPUT sub
Public ACTION_CHANNEL As String

'how many sockets to load on startup
Public Const maxtcp = 10

'is socket connected?
Public connected As Boolean

'your info
Public RealName  As String
Public email As String
Public nickname As String
Public server As String

'CHAT_Index = winsock index for DCC Chat's
Public CHAT_Index As Long
Public ChatWindow(1 To maxtcp) As New frmChat
Public ChatWindowName(1 To maxtcp) As String
Public ChatWindowx(1 To maxtcp) As New frmChat
Public ChatWindowNamex(1 To maxtcp) As String

'Notify Nickname list
'up to 100 nicknames
Public notify(1 To 100) As String
Public notifylist As String

'computer needs sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'right click
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_RBUTTONDOWN = &H204
'add image to list box
Public Const WM_SETREDRAW = &HB



'no refresh of the window
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'I dunno
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public PingReply As Long

'OK, script vars
Public Type numericcode
    NUM As String
    server As String
    nickname As String
    parms As String
    ServerText As String
End Type
Public Type commandtrigger
    parms As String
    username As String
    target As String
    command As String
    chanjoin As String
    chanpart As String
    nickjoin As String
    nickpart As String
End Type
Public Type Channelstats
    topic As String
    name As String
End Type
Public chanstats(1 To ChannelMax) As Channelstats
Public events As commandtrigger
Public raw As numericcode

'for dcc get files
Public FileIndex As Integer
'for dcc send files
Public FileListenPort As Integer

'for text strings in options
Public Type CustomStrings
    join As String
    part As String
    kick As String
    quit As String
    pm As String
End Type
Public strCustom As CustomStrings

Public colString As New Collection


'channel list variables
Public Type ChanListX
    max As Integer
    mim As Integer
    search As String
    
End Type

Public Sub CheckLine(strData As String)
    'On Error Resume Next
    Dim word() As String
    Dim parms As String
    Dim i As Integer
    'word = Split(strData, Chr(32))
    'Select Case LCase(word(0))
    '    Case "ping"
    '        mdiMain.tcp.SendData "PONG " & Mid(word(1), 2) & vbCrLf
    'End Select
    'If IsNumeric(word(1)) = True Then
    '    For i = 3 To UBound(word)
    '        parms = parms & word(i) & Chr(32)
    '    Next i
    '    parms = Mid(parms, 1, Len(parms) - 1)
    '    Call numeric(word(0), word(1), word(2), parms)
    'Else
    '    For i = 3 To UBound(word)
    '        parms = parms & word(i) & Chr(32)
    '    Next i
    '    parms = Mid(parms, 1, Len(parms) - 1)
    '    Call command(word(0), word(1), word(2), parms)
    'End If

End Sub
Public Sub CheckLine3(strLine As String)
    On Error Resume Next
    Dim i As Integer
    Dim oneline() As String
    Dim blComplete As Boolean
    blComplete = False
    Static RestLine As String
    
    If RestLine <> "" Then
        strLine = RestLine & strLine
        RestLine = ""
    End If
    
    If Right(strLine, 2) = vbCrLf Then
        blComplete = True
    End If
    
    oneline = Split(strLine, vbCrLf)
    
    If blComplete Then
        For i = 0 To UBound(oneline) - 1
            Call CheckWord(oneline(i))
            'DoEvents
        Next i
    Else
        RestLine = oneline(UBound(oneline) - 1)
        For i = 0 To UBound(oneline) - 2
            Call CheckWord(oneline(i))
        Next i
    End If
End Sub
Public Sub CheckLine2(strLine As String)
    On Error Resume Next
    Dim i As Integer
    Dim oneline() As String
    Static RestLine As String
    
    strLine = RestLine & strLine
    RestLine = ""
    
    'experimental word split char ‰
    strLine = Replace(strLine, vbCrLf, "‰")
    strLine = Replace(strLine, Chr(10), "‰")
    strLine = Replace(strLine, Chr(13), "‰")
    
    'this will see if the data was sent completly
    Dim LineComplete As Boolean
    LineComplete = False
    If Right(strLine, 1) = "‰" Then
        'line is complete
        LineComplete = True
        strLine = Mid(strLine, 1, Len(strLine) - 1)
    End If
    oneline = Split(strLine, "‰")
    'get last string and get it ready to attach to next string
    If LineComplete = False Then
        RestLine = oneline(UBound(oneline) - 1)
        For i = 0 To UBound(oneline) - 2
            Call CheckWord(oneline(i))
        Next i
    Else
        For i = 0 To UBound(oneline) - 1
            Call CheckWord(oneline(i))
        Next i
    End If

    Sleep 50
    
End Sub




Public Sub CheckWord2(strWord As String)
    raw.ServerText = strWord

    On Error Resume Next
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim word() As String
    Dim ErrorString As String
    If strWord = "" Then Exit Sub
    'check for carriage return in front of line and delete
    If Mid(strWord, 1, 1) = Chr(13) Or Mid(strWord, 1, 1) = Chr(10) Then
        strWord = Mid(strWord, 2)
    End If
    'this is used for ERRORs on connections
    'if the server is too full
    ErrorString = strWord
    'uncomment to show on status
    'frmStatus.txtStatus.SelText = "ONELINE: " & strWord & vbCrLf

    
    'split the commands into seperate words
    'ReDim Preserve statement is the KEY
    'Do Until InStr(strWord, Chr(32)) = 0
    Do Until InStr(strWord, Chr(32)) = 0
        X = InStr(strWord, Chr(32))
        If X Then
            Y = Y + 1
            ReDim Preserve word(Y)
            word(Y) = Mid(strWord, 1, X - 1)
            strWord = Mid(strWord, X + 1)
            'numeric?
        End If
        If Y = 3 Then
            If IsNumeric(word(2)) = True Then
                nickname = word(3)
                Call numeric(word(1), word(2), word(3), strWord)
                Call ClearVars
                Exit Sub
            End If
            Exit Do
        End If
     Loop
    ReDim Preserve word(Y + 1)
    word(Y + 1) = strWord

    Dim parms As String
    For i = 4 To UBound(word)
        If Trim(word(i)) <> "" Then
            parms = parms & " " & word(i)
        End If
    Next i
    If word(1) = "PING" Then
        mdiMain.tcp.SendData "PONG " & Mid(word(2), 2) & vbCrLf
        frmStatus.txtStatus.SelColor = RGB(0, 140, 0)
        frmStatus.txtStatus.SelText = "*** PONG " & Mid(word(2), 2) & vbCrLf
        'status window shows server name
        frmStatus.Caption = "Cabral Status: " & Mid(word(2), 2)
        'mdiMain shows [server name]
        server = Mid(word(2), 2)
        frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & Mid(word(2), 2) & ":" & mdiMain.tcp.RemotePort
        frmStatus.txtStatus.SelColor = vbBlack
        frmStatus.txtStatus.SelText = "-" & vbCrLf
        Exit Sub
    End If
    If word(1) = "ERROR:" Then
        DoColor frmStatus.txtStatus, "2*** " & ErrorString & vbCrLf & "-" & vbCrLf
    End If
    If word(1) = "ERROR" Then
        DoColor frmStatus.txtStatus, "2*** " & ErrorString & vbCrLf & "-" & vbCrLf
    End If
    Call command(word(1), word(2), word(3), parms)

    Call ClearVars
    'frmStatus.txtStatus.InsertContents SF_TEXT, "WORD 2: " & word(2) & " " & word(3) & vbCrLf
    'frmStatus.txtStatus.InsertContents SF_TEXT, "PARMS: " & parms & vbCrLf
    
    'script
End Sub
Function ReplaceX(xString As String, xOldWord As String, xNewWord As String) As String
    Dim i As Integer
    Dim part1, part2 As String
    Do Until InStr(xString, xOldWord) = 0
        i = InStr(xString, xOldWord)
        If i Then
            part1 = Mid(xString, 1, i - 1)
            part2 = Mid(xString, i + Len(xOldWord))
            xString = part1 & xNewWord & part2
        End If
    Loop
    ReplaceX = xString
End Function

Public Sub AddTaskbar(xCaption As String, picType As Integer)
    mdiMain.StatusBar.Panels.Add (mdiMain.StatusBar.Panels.Count + 1), xCaption, xCaption
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).MinWidth = 500
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).AutoSize = sbrSpring
    'mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).AutoSize = sbrContents
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Bevel = sbrRaised 'sbrInset
    mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Style = sbrText
    'mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).AutoSize = sbrSpring
    
    Select Case picType
        Case 1 'CHANNEL
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(4).Picture
        Case 2
            'black box
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(2).Picture
        Case 3
            '@ box
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(3).Picture
        Case 4
            'Query msg
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(4).Picture
        Case 5
            'Status/Client Msg
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(5).Picture
        Case 6
            'CTCP
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(6).Picture
        Case 7
            'MOTD
            mdiMain.StatusBar.Panels.Item((mdiMain.StatusBar.Panels.Count)).Picture = mdiMain.imgTaskbar.ListImages(7).Picture
    End Select
End Sub
Public Sub RemoveTaskbar(xCaption As String)
    Dim i As Integer
    
    For i = 1 To mdiMain.StatusBar.Panels.Count
        'MsgBox mdiMain.StatusBar.Panels.Item(i).Key & " --- " & xCaption
        If LCase(mdiMain.StatusBar.Panels.Item(i).Key) = LCase(xCaption) Then
            mdiMain.StatusBar.Panels.Remove (i)
            Exit For
        End If
    Next i
End Sub




Function FileExists(xFilename As String) As Boolean
    Dim results As String
    results = Dir$(xFilename)
    
    If results = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Public Function xSplit(SplitString As String, SplitLetter As String) As Variant
    ReDim SplitArray(0 To 0) As Variant
    Dim TempLetter As String
    Dim TempSplit As String
    Dim i As Integer
    Dim X As Integer
    Dim StartPos As Integer
    
    X = 0
    
    SplitString = SplitString & SplitLetter


    For i = 1 To Len(SplitString)
        TempLetter = Mid(SplitString, i, Len(SplitLetter))


        If TempLetter = SplitLetter Then
            TempSplit = Mid(SplitString, (StartPos + 1), (i - StartPos) - 1)


            If TempSplit <> "" Then
                ReDim Preserve SplitArray(0 To X) As Variant
                SplitArray(X) = TempSplit
                X = X + 1
            End If
            StartPos = i
        End If
    Next i
    'Split = SplitArray
End Function
Public Function CheckListbox(strListBox As ListBox, CheckName As String) As Boolean
    Dim i As Integer
    CheckListbox = False
    For i = 0 To strListBox.ListCount - 1
        If strListBox.List(i) = CheckName Then
            CheckListbox = True
            Exit For
        End If
    Next i
End Function

Public Sub ShowStats(RTF As RichTextBox)
    Dim ops As Integer
    Dim voiced As Integer
    Dim users As Integer
    Dim i As Integer
    Dim X As Integer
    'lets find channel
    For i = 1 To ChannelMax
        If LCase(ACTION_CHANNEL) = LCase(ChannelName(i)) Then
            For X = 0 To channel(i).lstNames.ListCount - 1
                Select Case left(channel(i).lstNames.List(X), 1)
                    Case "@"
                        ops = ops + 1
                    Case "+"
                        voiced = voiced + 1
                    Case Else
                        users = users + 1
                End Select
            Next X
            DoColor channel(i).txtText, "" & color.join & "*** Stats: Ops(" & ops & ") Voiced(" & voiced & ") Users(" & users & ") - Total:" & ops + voiced + users
            Exit For
        End If
    Next i
    
End Sub
Public Sub ClearVars()
    raw.nickname = ""
    raw.NUM = ""
    raw.parms = ""
    raw.server = ""
    raw.ServerText = ""
    events.username = ""
    events.target = ""
    events.parms = ""
    events.nickpart = ""
    events.nickjoin = ""
    events.command = ""
    events.chanpart = ""
    events.chanjoin = ""
End Sub

Public Sub CheckWord(strWord As String)
    raw.ServerText = strWord

    On Error Resume Next
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim word() As String
    Dim ErrorString As String
    Dim parms As String
    If strWord = "" Then Exit Sub
    
    'this is used for ERRORs on connections
    'if the server is too full
    ErrorString = strWord

    
    'split the commands into seperate words
    word = Split(strWord, Chr(32))
    
    'ReDim Preserve statement is the KEY
    'Do Until InStr(strWord, Chr(32)) = 0
    
    For i = 3 To UBound(word)
        If Trim(word(i)) <> "" Then
            parms = parms & " " & word(i)
        End If
    Next i
    
    If IsNumeric(word(1)) = True Then
        nickname = word(2)
        Call numeric(word(0), word(1), word(2), Mid(parms, 2))
        Exit Sub
    End If

    If word(0) = "PING" Then
        Call DoColor(frmStatus.txtStatus, "3*** PONG from " & Mid(word(1), 2) & "!" & vbCrLf & "-")
        mdiMain.tcp.SendData "PONG " & Mid(word(1), 2) & vbCrLf
        'mdiMain shows [server name]
        server = Mid(word(1), 2)
        frmStatus.Caption = "Status: [" & MyModes & "] " & nickname & " on " & Mid(word(1), 2) & ":" & mdiMain.tcp.RemotePort
        Exit Sub
    End If
    
    Call command(word(0), word(1), word(2), Mid(parms, 2))

End Sub

Sub UserOn(user_query As RichTextBox, StrName As String)
    Dim i As Integer, X As Integer
    For i = 1 To ChannelMax
        If ChannelName(i) <> "" Then
            For X = 0 To channel(i).lstNames.ListCount - 1
                If LCase(channel(i).lstNames.List(X)) = LCase(StrName) Or LCase(channel(i).lstNames.List(X)) = LCase("@" & StrName) Or LCase(channel(i).lstNames.List(X)) = LCase("+" & StrName) Then
                    DoColor user_query, "" & color.notice & "* " & StrName & " is on " & ChannelName(i)
                End If
            Next X
        End If
    Next i
End Sub
Sub UpdateCaption(Index As Integer)
    If ChannelLimit(Index) = "" Then
        channel(Index).Caption = ChannelName(Index) & " [" & channel(Index).lstNames.ListCount & "] [+" & ChannelModes(Index) & "] :" & ChannelTopic(Index)
    Else
        channel(Index).Caption = ChannelName(Index) & " [" & channel(Index).lstNames.ListCount & "] [+" & ChannelModes(Index) & " " & ChannelLimit(Index) & "] :" & ChannelTopic(Index)
    End If
End Sub
Public Sub RemoveNode(StrName As String)
    Dim i As Integer
    For i = 1 To mdiMain.tvMain.Nodes.Count
        'DoColor frmStatus.txtStatus, "4...." & i & LCase(mdiMain.tvMain.Nodes.Item(i).Text)
        If LCase(StrName) = LCase(mdiMain.tvMain.Nodes.Item(i).Text) Then
            mdiMain.tvMain.Nodes.Remove i
            Exit Sub
        End If
    Next i
End Sub
Public Sub highlight_node(searchword As String)
    On Error Resume Next
    Dim i As Integer
    Dim strServer() As String
    
    For i = 1 To mdiMain.tvMain.Nodes.Count
        If LCase(searchword) = LCase(mdiMain.tvMain.Nodes.Item(i).Text) Then
            mdiMain.tvMain.Nodes.Item(i).ForeColor = vbRed
        End If
    Next i
End Sub
Public Sub UNhighlight_node(searchword As String)
    On Error Resume Next
    Dim i As Integer
    
    For i = 1 To mdiMain.tvMain.Nodes.Count
        If LCase(searchword) = LCase(mdiMain.tvMain.Nodes.Item(i).Text) Then
            mdiMain.tvMain.Nodes.Item(i).ForeColor = vbBlack
        End If
    Next i
End Sub
Public Sub addURL(strURL As String, User As String, timestamp As String)

'    Open App.Path & "\url.txt" For Output As #1
    Open App.Path & "\url.txt" For Append As #1
        Print #1, strURL & Chr(1) & User & Chr(1) & timestamp
    Close #1

End Sub
