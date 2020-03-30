Attribute VB_Name = "Ping"
'Unknow Creator
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As ip_option_information, ReplyBuffer As icmp_echo_reply, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Const PING_TIMEOUT = 1200
Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYSSTATUS_LEN = 256
Public Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
Public Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
Public Const SOCKET_ERROR = -1
Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequested As Integer, lpWSAData As tagWSAData) As Integer
Public Declare Function WSACleanup Lib "wsock32" () As Integer

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (IP_STATUS_BASE + 1)
Public Const IP_DEST_NET_UNREACHABLE = (IP_STATUS_BASE + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (IP_STATUS_BASE + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (IP_STATUS_BASE + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (IP_STATUS_BASE + 5)
Public Const IP_NO_RESOURCES = (IP_STATUS_BASE + 6)
Public Const IP_BAD_OPTION = (IP_STATUS_BASE + 7)
Public Const IP_HW_ERROR = (IP_STATUS_BASE + 8)
Public Const IP_PACKET_TOO_BIG = (IP_STATUS_BASE + 9)
Public Const IP_REQ_TIMED_OUT = (IP_STATUS_BASE + 10)
Public Const IP_BAD_REQ = (IP_STATUS_BASE + 11)
Public Const IP_BAD_ROUTE = (IP_STATUS_BASE + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (IP_STATUS_BASE + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (IP_STATUS_BASE + 14)
Public Const IP_PARAM_PROBLEM = (IP_STATUS_BASE + 15)
Public Const IP_SOURCE_QUENCH = (IP_STATUS_BASE + 16)
Public Const IP_OPTION_TOO_BIG = (IP_STATUS_BASE + 17)
Public Const IP_BAD_DESTINATION = (IP_STATUS_BASE + 18)
Public Const IP_ADDR_DELETED = (IP_STATUS_BASE + 19)
Public Const IP_SPEC_MTU_CHANGE = (IP_STATUS_BASE + 20)
Public Const IP_MTU_CHANGE = (IP_STATUS_BASE + 21)
Public Const IP_UNLOAD = (IP_STATUS_BASE + 22)
Public Const IP_ADDR_ADDED = (IP_STATUS_BASE + 23)
Public Const IP_GENERAL_FAILURE = (IP_STATUS_BASE + 50)
Public Const MAX_IP_STATUS = IP_STATUS_BASE + 50
Public Const IP_PENDING = (IP_STATUS_BASE + 255)

Public Type ip_option_information
 TTL             As Byte     'Time To Live [ Nb Max de sauts de routeurs ]
 Tos             As Byte     'Type Of Service [ Type de trame ]
 flags           As Byte     'IP header flags [ En-tête de la trame ]
 OptionsSize     As Byte     'Taille des trames
 OptionsData     As Long     'Options ( hops, TargetIP,... )
End Type

Public Type icmp_echo_reply
 Address         As Long                     'Replying address
 Status          As Long                     'Reply IP_STATUS, values as defined above
 RoundTripTime   As Long                     'RTT in milliseconds
 DataSize        As Integer                  'Reply data size in bytes
 Reserved        As Integer                  'Reserved for system use
 DataPointer     As Long                     'Pointer to the reply data
 Options         As ip_option_information    'Reply options
 Data            As String * 250             'Reply data which should be a copy of the string sent, NULL terminated
End Type                                     ' this field length should be large enough to contain the string sent

Public Type tagWSAData
 wVersion            As Integer
 wHighVersion        As Integer
 szDescription       As String * WSADESCRIPTION_LEN_1
 szSystemStatus      As String * WSASYSSTATUS_LEN_1
 iMaxSockets         As Integer
 iMaxUdpDg           As Integer
 lpVendorInfo        As String * 200
End Type

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Public ReturnedRoundTime$
Public ReturnedIP$
Public ReturnedTTL$
Public LoopPing As Boolean
Public Ancre As Boolean
Public SelectedIp$


Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function PingIp(ByVal IP As String, TTL As Integer, Timeout As Integer)
 '   Va retourner  l'ip du routeur de distance TTL
 ' Si le TTL est suffisant pour atteindre la cible
 ' alors  on retourne 1.  Sinon on retourne 0 pour
 ' indiquer  qu'un nouveau  routeur à été atteind.
 ' Si une erreure se produit, on retourne -1
 
 Dim hFile       As Long
 Dim lRet        As Long
 Dim lIPAddress  As Long
 Dim strMessage  As String
 Dim pOptions    As ip_option_information
 Dim pReturn     As icmp_echo_reply
 Dim iVal        As Integer
 Dim lPingRet    As Long
 Dim pWsaData    As tagWSAData
 Dim inicio As Long
' Statusmain.lblStatus.Caption = "Checking..."
inicio = GetTickCount
 On Error Resume Next
 strMessage = ""
 iVal = WSAStartup(&H101, pWsaData)
 lIPAddress = ConvertIPAddressToLong(IP)
 hFile = IcmpCreateFile()
    
 pOptions.TTL = TTL
    
 lRet = IcmpSendEcho(hFile, lIPAddress, strMessage, 10, pOptions, pReturn, Len(pReturn), Timeout)
 DoEvents
 
inicio = GetTickCount - inicio
 If lRet = 0 Then
  PingIp = -1
 Else
  If pReturn.Status <> 0 Then
   PingIp = inicio
   ReturnedTTL$ = Str$(pReturn.Options.TTL)
   ReturnedRoundTime$ = Str$(pReturn.RoundTripTime)
   ReturnedIP$ = LongToIp$(Hex$(pReturn.Address))
  Else
   PingIp = inicio
   ReturnedTTL$ = Str$(pReturn.Options.TTL)
   ReturnedRoundTime$ = Str$(pReturn.RoundTripTime)
   ReturnedIP$ = LongToIp$(Hex$(pReturn.Address))
  End If
 End If
 
 lRet = IcmpCloseHandle(hFile)
 iVal = WSACleanup()
End Function

Function ConvertIPAddressToLong(strAddress As String) As Long
 ' Convertion chaine IP en long
 
 Dim strTemp             As String
 Dim lAddress            As Long
 Dim iValCount           As Integer
 Dim lDotValues(1 To 4)  As String
    
 strTemp = strAddress
 iValCount = 0
    
 While InStr(strTemp, ".") > 0
  iValCount = iValCount + 1
  lDotValues(iValCount) = Mid(strTemp, 1, InStr(strTemp, ".") - 1)
  strTemp = Mid(strTemp, InStr(strTemp, ".") + 1)
 Wend
        
 iValCount = iValCount + 1
 lDotValues(iValCount) = strTemp
    
 If iValCount <> 4 Then
  ConvertIPAddressToLong = 0
  Exit Function
 End If
        
 lAddress = Val("&H" & Right("00" & Hex(lDotValues(4)), 2) & Right("00" & Hex(lDotValues(3)), 2) & Right("00" & Hex(lDotValues(2)), 2) & Right("00" & Hex(lDotValues(1)), 2))
               
 ConvertIPAddressToLong = lAddress
End Function

Function LongToIp$(Value$)
 ' Convertion d'un long en addresse IP ( chaine )
 
 Value$ = "00000" + Value$
 Value$ = Right$(Value$, 8)
 op1$ = Right$(Value$, 2)
 op2$ = Mid$(Value$, 5, 2)
 op3$ = Mid$(Value$, 3, 2)
 op4$ = Left$(Value$, 2)
 LongToIp$ = HexDec$(op1$) + "." + HexDec$(op2$) + "." + HexDec$(op3$) + "." + HexDec$(op4$)
End Function

Function HexDec$(Value$)
 ' Convertion Hexa en Decimal
 
 Id = 0: Result = 0
 For i = Len(Value$) To 1 Step -1
  If Mid$(Value$, i, 1) = "0" Then Vl = 0
  If Mid$(Value$, i, 1) = "1" Then Vl = 1
  If Mid$(Value$, i, 1) = "2" Then Vl = 2
  If Mid$(Value$, i, 1) = "3" Then Vl = 3
  If Mid$(Value$, i, 1) = "4" Then Vl = 4
  If Mid$(Value$, i, 1) = "5" Then Vl = 5
  If Mid$(Value$, i, 1) = "6" Then Vl = 6
  If Mid$(Value$, i, 1) = "7" Then Vl = 7
  If Mid$(Value$, i, 1) = "8" Then Vl = 8
  If Mid$(Value$, i, 1) = "9" Then Vl = 9
  If Mid$(Value$, i, 1) = "A" Then Vl = 10
  If Mid$(Value$, i, 1) = "B" Then Vl = 11
  If Mid$(Value$, i, 1) = "C" Then Vl = 12
  If Mid$(Value$, i, 1) = "D" Then Vl = 13
  If Mid$(Value$, i, 1) = "E" Then Vl = 14
  If Mid$(Value$, i, 1) = "F" Then Vl = 15
  Result = Result + (Vl * 16 ^ Id)
  Id = Id + 1
 Next i
 HexDec$ = Str$(Result)
 HexDec$ = LTrim$(HexDec$)
End Function
