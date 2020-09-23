Attribute VB_Name = "irc_input"
Option Explicit
'ident public variables
Public IdentUserID As String

Public dns As New dns

'IRC OPTIONS
Public Type SHOW
    quits As Integer
    joinpart As Integer
    modes As Integer
    topics As Integer
    kicks As Integer
    motd As Integer
    channelfolder As Integer
    address As Integer
    whoisnotify As Integer
    notifylist As Integer
End Type
Public iShow As SHOW

Public Type LoginModes
    i As Integer
    w As Integer
    s As Integer
End Type
Public iModes As LoginModes

'INI api
Public iniFilename As String
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section, KeyName As String) As String
    Dim sRet As String
    Dim FileName As String
    FileName = App.Path & "\" & iniFilename
    
    sRet = String(255, Chr(0))
    ReadINI = left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String) As Integer
    Dim r
    Dim sFilename As String
    sFilename = App.Path & "\" & iniFilename
    
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFilename)
End Function

Sub xINPUT(strData As String, ByVal RTF As RichTextBox)
    
    strData = RTrim(strData)
    Dim X, Y As Integer
    'don't set array it'll be set with ReDim
    Dim word() As String
    Dim parms As String
    
    'used for dns requests
    Dim retDNS As String

    'split the commands into seperate words
    word = Split(strData, " ")
        'for some you'll need parameters
        For X = 1 To UBound(word)
            parms = parms & word(X) & Chr(32)
        Next X
        parms = RTrim(parms)
        'DO IT
        
        Select Case UCase(word(0))
            Case "RAW"
                mdiMain.tcp.SendData parms & vbCrLf
                frmStatus.txtStatus.SelText = "->Server: " & parms & vbCrLf
                frmStatus.txtStatus.SelText = "-" & vbCrLf
            Case "MSG"
                parms = ""
                For X = 2 To UBound(word)
                    parms = parms & word(X) & Chr(32)
                Next X
               
                parms = RTrim(parms)
                mdiMain.tcp.SendData "PRIVMSG " & word(1) & " :" & parms & vbCrLf
                Call DoColor(frmStatus.txtStatus, "-> *" & word(1) & "* " & parms)
                frmStatus.txtStatus.SelText = "-" & vbCrLf
            Case "WHOIS"
                mdiMain.tcp.SendData "WHOIS " & word(1) & vbCrLf
            Case "JOIN"
                If left(word(1), 1) = "#" Then
                    mdiMain.tcp.SendData "JOIN " & word(1) & vbCrLf
                Else
                    mdiMain.tcp.SendData "JOIN #" & word(1) & vbCrLf
                End If
            Case "PART"
                If left(word(1), 1) = "#" Then
                    mdiMain.tcp.SendData "PART " & word(1) & vbCrLf
                Else
                    mdiMain.tcp.SendData "PART #" & word(1) & vbCrLf
                End If
            Case "NICK"
                mdiMain.tcp.SendData "NICK " & word(1) & vbCrLf
            Case "CHAT"
                Dim LIP As String
                For X = 1 To 25
                    If mdiMain.CHATx(X).State = sckClosed Or mdiMain.CHATx(X).State = sckError Then
                        mdiMain.CHATx(X).Close
                        mdiMain.CHATx(X).LocalPort = mdiMain.CHATx(0).LocalPort
                        mdiMain.CHATx(X).Listen
                        DoColor frmStatus.txtStatus, "4* Requesting chat with " & word(1) & "   [" & mdiMain.CHATx(X).LocalPort & "]"
                        Load ChatWindowx(X)
                        ChatWindowx(X).SHOW
                        ChatWindowx(X).Caption = word(1) '& " - " & mdiMain.CHATx(X).LocalPort
                        ChatWindowNamex(X) = word(1) '& " - " & mdiMain.CHATx(X).LocalPort
                        DoColor ChatWindowx(X).txtDCC, "4* Waiting for acknowledgement... "
                        mdiMain.tcp.SendData "NOTICE " & word(1) & " :DCC CHAT (" & mdiMain.tcp.LocalIP & ")" & vbCrLf
                        LIP = IrcGetLongIP(mdiMain.tcp.LocalIP)
                        mdiMain.tcp.SendData "PRIVMSG " & word(1) & " :DCC CHAT chat " & LIP & " " & mdiMain.CHATx(X).LocalPort & "" & vbCrLf
                        Exit For
                    End If
                Next X
            Case "NAMES"
                If word(1) <> "" Then
                    mdiMain.tcp.SendData "NAMES " & word(1) & vbCrLf
                End If
            Case "QUIT"
                mdiMain.tcp.SendData "QUIT :" & parms & vbCrLf
            Case "ME"
                DoColor RTF, "" & color.action & "*" & nickname & " " & parms
                mdiMain.tcp.SendData "PRIVMSG " & ACTION_CHANNEL & " :ACTION " & parms & "" & vbCrLf
            Case "CLEAR"
                RTF.Text = ""
            Case "LIST"
                If parms <> "" Then
                    mdiMain.tcp.SendData "LIST " & word(1) & vbCrLf
                Else
                    mdiMain.tcp.SendData "LIST" & vbCrLf
                End If
            Case "MOTD"
                mdiMain.tcp.SendData "MOTD" & vbCrLf
            Case "LUSERS"
                mdiMain.tcp.SendData "LUSERS" & vbCrLf
            Case "STATS"
                ShowStats RTF
            Case "ECHO"
                DoColor RTF, "" & color.normal & parms
            Case "DNSNAME"
                DoColor RTF, "" & color.action & "*** Looking up " & word(1) & vbCrLf & "-"
                retDNS = dns.NameToAddress(word(1))
                DoColor RTF, "" & color.action & "*** Resolved " & word(1) & " to " & retDNS & vbCrLf & "-"
            Case "DNSIP"
                DoColor RTF, "" & color.action & "*** Looking up " & word(1) & vbCrLf & "-"
                retDNS = dns.AddressToName(word(1))
                DoColor RTF, "" & color.action & "*** Resolved " & word(1) & " to " & retDNS & vbCrLf & "-"
            Case "DNS"
                DoColor RTF, "" & color.action & "*** Looking up " & word(1) & vbCrLf & "-"
                If IsNumeric(left(word(1), 1)) Then
                    retDNS = dns.AddressToName(word(1))
                Else
                    retDNS = dns.NameToAddress(word(1))
                End If
                DoColor RTF, "" & color.action & "*** Resolved " & word(1) & " to " & retDNS & vbCrLf & "-"
            Case "PING"
                mdiMain.tcp.SendData "PRIVMSG " & word(1) & " :" & Chr$(1) & "PING " & Trim(Str(DateDiff("s", CVDate("01/01/1970"), Now))) & Chr$(1) & vbCrLf
            Case "REFRESHLIST"
                frmListChannels.lstChannels.Refresh
            Case Else
                frmStatus.txtStatus.SelText = "->Server: " & strData & vbCrLf & "-" & vbCrLf
                mdiMain.tcp.SendData strData & vbCrLf
        End Select

End Sub

