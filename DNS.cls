VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbInitialized As Boolean

Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128

Const AF_INET = 4&

Private Type HOSTENT
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Type WSADATA
  wversion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpszVendorInfo As Long
End Type

Private Declare Function WSAStartup _
                                Lib "wsock32" _
                                (ByVal VersionReq As Long, _
                                WSADataReturn As WSADATA) _
                                As Long

Private Declare Function WSACleanup _
                                Lib "wsock32" _
                                () _
                                As Long

Private Declare Function WSAGetLastError _
                                Lib "wsock32" _
                                () _
                                As Long

Private Declare Function gethostbyaddr _
                                Lib "wsock32" _
                                (addr As Long, _
                                addrLen As Long, _
                                addrType As Long) _
                                As Long

Private Declare Function gethostbyname _
                                Lib "wsock32" _
                                (ByVal hostname As String) _
                                As Long

Private Declare Sub RtlMoveMemory _
                                Lib "kernel32" _
                                (hpvDest As Any, _
                                ByVal hpvSource As Long, _
                                ByVal cbCopy As Long)

'checks if string is valid IP address
Private Function CheckIP(IPToCheck As String) As Boolean

  Dim TempValues
  Dim iLoop As Long
  Dim TempByte As Byte
  
  On Error GoTo CheckIPError
  
  TempValues = Split(IPToCheck, ".")
  
  If UBound(TempValues) < 3 Then
    Exit Function
  End If
  
  For iLoop = LBound(TempValues) To UBound(TempValues)
    TempByte = TempValues(iLoop)
  Next iLoop
  CheckIP = True
  
CheckIPError:

End Function

'converts IP address from string to sin_addr
Private Function MakeIP(strIP As String) As Long
    
  Dim vTemp
  Dim lngTemp As Long
  Dim iLoop As Long
  
  On Error GoTo MakeIPError
  
  vTemp = Split(strIP, ".")
  
  For iLoop = 0 To (UBound(vTemp) - 1)
    lngTemp = lngTemp + (vTemp(iLoop) * (256 ^ iLoop))
  Next iLoop
  
  If vTemp(UBound(vTemp)) < 128 Then
    lngTemp = lngTemp + (vTemp(UBound(vTemp)) * (256 ^ 3))
  Else
    lngTemp = lngTemp + ((vTemp(UBound(vTemp)) - 256) * (256 ^ 3))
  End If
  
  MakeIP = lngTemp
MakeIPError:
End Function

'resolves IP address to host name
Public Function AddrToName(strAddr As String) As String
  
  Dim heEntry As HOSTENT
  Dim strHost As String * 255
  Dim strTemp As String
  Dim lngRet As Long
  Dim lngIP As Long
  
  On Error GoTo AddrToNameError
  
  If CheckIP(strAddr) Then
    lngIP = MakeIP(strAddr)
    lngRet = gethostbyaddr(lngIP, 4, AF_INET)
    If lngRet = 0 Then
      Exit Function
    End If
    RtlMoveMemory heEntry, lngRet, Len(heEntry)
    RtlMoveMemory ByVal strHost, heEntry.hName, 255
    strTemp = TrimNull(strHost)
    AddrToName = strTemp
  End If

AddrToNameError:
End Function
'resolves host name to IP address
Public Function NameToAddr(ByVal strHost As String)
  
  Dim ip_list() As Byte
  Dim heEntry As HOSTENT
  Dim strIPAddr As String
  Dim lp_HostEnt As Long
  Dim lp_HostIP As Long
  Dim iLoop As Integer
  
  On Error GoTo NameToAddrError
  
  lp_HostEnt = gethostbyname(strHost)
  
  If lp_HostEnt = 0 Then
    Exit Function
  End If
  
  RtlMoveMemory heEntry, lp_HostEnt, LenB(heEntry)
  RtlMoveMemory lp_HostIP, heEntry.hAddrList, 4
  
  ReDim ip_list(1 To heEntry.hLength)
  
  RtlMoveMemory ip_list(1), lp_HostIP, heEntry.hLength
  
  For iLoop = 1 To heEntry.hLength
    strIPAddr = strIPAddr & ip_list(iLoop) & "."
  Next
  
  strIPAddr = Mid(strIPAddr, 1, Len(strIPAddr) - 1)
  
  NameToAddr = strIPAddr
NameToAddrError:
  
End Function
Public Function AddressToName(strIP As String) As String
  
  If mbInitialized Then
    AddressToName = AddrToName(strIP)
  End If

End Function

Public Function NameToAddress(StrName As String) As String
  
  If mbInitialized Then
    NameToAddress = NameToAddr(StrName)
  End If

End Function

Private Function TrimNull(sTrim As String) As String

  Dim iFind As Long

  iFind = InStr(1, sTrim, Chr(0))
  If iFind > 0 Then
    TrimNull = left(sTrim, iFind - 1)
  Else
    TrimNull = sTrim
  End If

End Function

Private Sub Class_Initialize()

  Dim wsa As WSADATA
  mbInitialized = (WSAStartup(257, wsa) = 0)
    
End Sub

Private Sub Class_Terminate()

  If mbInitialized Then
    WSACleanup
  End If

End Sub


