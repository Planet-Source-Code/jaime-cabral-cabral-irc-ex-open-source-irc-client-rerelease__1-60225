Attribute VB_Name = "time"
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Function GetGMTBias() As Integer
    Dim lpTimeZoneInformation As TIME_ZONE_INFORMATION
    
    GetTimeZoneInformation lpTimeZoneInformation
    GetGMTBias = lpTimeZoneInformation.Bias
End Function


Public Function GetGMTBiasString() As String
    Dim X As Long, Y As Long
    X = -GetGMTBias
    Y = X Mod 60
   
    X = X \ 60
    If Y < 0 Then
        Y = -Y
        GetGMTBiasString = "GMT-" & Format$(X, "00") & ":" & _
        Format$(Y, "00")
    ElseIf X < 0 Then
        GetGMTBiasString = "GMT-" & _
        Format$(X, "00") & ":" & Format$(Y, "00")
    Else
        GetGMTBiasString = "GMT+" & _
        Format$(X, "00") & ":" & Format$(Y, "00")
    End If
End Function


Function CTime() As Long
    CTime = toCTime(Now)
End Function

Function toCTime(d As Date) As Long
    toCTime = DateDiff("s", CDate(#1/1/1970# - GetGMTBias / 60 / 24), d)
End Function

Function AscTime(CTime As Long) As Date
    AscTime = CDate(#1/1/1970# - GetGMTBias / 60 / _
    24) + (CTime / 3600& / 24)
End Function
Public Function irc_time(TheTime As String) As String
    Dim lpTime As TIME_ZONE_INFORMATION
    Dim TheDate As Date
    Dim Msg
    
    On Error GoTo err_handle
    
    TheDate = "January 1 1970 00:00:00"
    GetTimeZoneInformation lpTime
    
    If IsNumeric(TheTime) Then
        'given timer
        irc_time = Format(DateAdd("s", Val(TheTime) - (lpTime.Bias * 60), TheDate), "ddd mmm dd yyyy hh:mm:ss")
    Else
        'given date
        irc_time = (DateDiff("s", TheDate, TheTime) - (lpTime.Bias * 60))
    End If
    Exit Function
err_handle:
    irc_time = "Invalid date/time format"
End Function



