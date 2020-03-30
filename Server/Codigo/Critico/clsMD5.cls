VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsMD5.cls - Computes MD5 hashes of strings and files alike
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 17/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************


'**************************************************************
'                       N   O   T   E
'**************************************************************
'The MD5 algorithm makes great use of 4 advantages of C:
' - strings and byte arrays are the same
' - overflow is never checked
' - there exist shifting operations (<< and >>)
' - variables can be unsigned
'
'As you may know none of these are supported by VB, so I had to code
'several auxiliar methods to replace them.
'
' - Shifts were faked with the ShiftLeft and ShiftRight methods.
' - Overflow was tricked by preventing the first bit from being set
'in any way except an Or operation.
' - As for byte arrays - strings I had to use non-constant size byte arrays
'and StrConv() to change from and to Unicode
' - Unsigned longs were replaced with Currency. However, since bitwise
'operations and Mod are not defined for anything larger than a Long I had
'to do some conversions back and forth all over the code. For worst,
'the bit configuration of Currency and longs are completely different,
'therefore I couldn't simply move memory blocks, but had to abuse of
'maths operations.
'You will find 2 auxiliar methods:
'   * CurToLong takes a cirrency in the range of an unsigned long
'(if it existed in VB) and returns a long with the bit configuration
'such unsigned long would have (may be negative, but since the point
'is to use it for bitwise operations it doesn't mattter).
'   * LongToCur takes a long and returns a currency with the value
'and unsigned long would have if it had the bit configuration presented
'by the passed long.
'
'I know it's hard to get and really weird. Moreover, some maths
'may not be completely clear or the most efficient, however
'you can trust me it works perfectly, and as long as you don't abuse
'of the algorithm you should have no issues
'**************************************************************

Option Explicit

'Extreme values of 32 bit vars (used to avoid overflows)
Private Const MAX_ULONG As Currency = 2 ^ 32 - 1
Private Const MIN_LONG As Currency = -(2 ^ 31)
Private Const MAX_LONG As Currency = 2 ^ 31 - 1

' Magic initialization constants
Private Const MD5_INIT_STATE_0 As Currency = 1732584193
Private Const MD5_INIT_STATE_1 As Currency = 4023233417#
Private Const MD5_INIT_STATE_2 As Currency = 2562383102#
Private Const MD5_INIT_STATE_3 As Currency = 271733878

' Constants for Transform routine.
Private Const MD5_S11 As Byte = 7
Private Const MD5_S12 As Byte = 12
Private Const MD5_S13 As Byte = 17
Private Const MD5_S14 As Byte = 22
Private Const MD5_S21 As Byte = 5
Private Const MD5_S22 As Byte = 9
Private Const MD5_S23 As Byte = 14
Private Const MD5_S24 As Byte = 20
Private Const MD5_S31 As Byte = 4
Private Const MD5_S32 As Byte = 11
Private Const MD5_S33 As Byte = 16
Private Const MD5_S34 As Byte = 23
Private Const MD5_S41 As Byte = 6
Private Const MD5_S42 As Byte = 10
Private Const MD5_S43 As Byte = 15
Private Const MD5_S44 As Byte = 21

' Transformation Constants - Round 1
Private Const MD5_T01 As Currency = 3614090360# 'Transformation Constant 1
Private Const MD5_T02 As Currency = 3905402710# 'Transformation Constant 2
Private Const MD5_T03 As Currency = 606105819 'Transformation Constant 3
Private Const MD5_T04 As Currency = 3250441966# 'Transformation Constant 4
Private Const MD5_T05 As Currency = 4118548399# 'Transformation Constant 5
Private Const MD5_T06 As Currency = 1200080426 'Transformation Constant 6
Private Const MD5_T07 As Currency = 2821735955# 'Transformation Constant 7
Private Const MD5_T08 As Currency = 4249261313# 'Transformation Constant 8
Private Const MD5_T09 As Currency = 1770035416 'Transformation Constant 9
Private Const MD5_T10 As Currency = 2336552879# 'Transformation Constant 10
Private Const MD5_T11 As Currency = 4294925233# 'Transformation Constant 11
Private Const MD5_T12 As Currency = 2304563134# 'Transformation Constant 12
Private Const MD5_T13 As Currency = 1804603682 'Transformation Constant 13
Private Const MD5_T14 As Currency = 4254626195# 'Transformation Constant 14
Private Const MD5_T15 As Currency = 2792965006# 'Transformation Constant 15
Private Const MD5_T16 As Currency = 1236535329 'Transformation Constant 16

'Transformation Constants - Round 2
Private Const MD5_T17 As Currency = 4129170786# 'Transformation Constant 17
Private Const MD5_T18 As Currency = 3225465664# 'Transformation Constant 18
Private Const MD5_T19 As Currency = 643717713 'Transformation Constant 19
Private Const MD5_T20 As Currency = 3921069994# 'Transformation Constant 20
Private Const MD5_T21 As Currency = 3593408605# 'Transformation Constant 21
Private Const MD5_T22 As Currency = 38016083  'Transformation Constant 22
Private Const MD5_T23 As Currency = 3634488961# 'Transformation Constant 23
Private Const MD5_T24 As Currency = 3889429448# 'Transformation Constant 24
Private Const MD5_T25 As Currency = 568446438 'Transformation Constant 25
Private Const MD5_T26 As Currency = 3275163606# 'Transformation Constant 26
Private Const MD5_T27 As Currency = 4107603335# 'Transformation Constant 27
Private Const MD5_T28 As Currency = 1163531501 'Transformation Constant 28
Private Const MD5_T29 As Currency = 2850285829# 'Transformation Constant 29
Private Const MD5_T30 As Currency = 4243563512# 'Transformation Constant 30
Private Const MD5_T31 As Currency = 1735328473 'Transformation Constant 31
Private Const MD5_T32 As Currency = 2368359562# 'Transformation Constant 32

'Transformation Constants - Round 3
Private Const MD5_T33 As Currency = 4294588738# 'Transformation Constant 33
Private Const MD5_T34 As Currency = 2272392833# 'Transformation Constant 34
Private Const MD5_T35 As Currency = 1839030562 'Transformation Constant 35
Private Const MD5_T36 As Currency = 4259657740# 'Transformation Constant 36
Private Const MD5_T37 As Currency = 2763975236# 'Transformation Constant 37
Private Const MD5_T38 As Currency = 1272893353 'Transformation Constant 38
Private Const MD5_T39 As Currency = 4139469664# 'Transformation Constant 39
Private Const MD5_T40 As Currency = 3200236656# 'Transformation Constant 40
Private Const MD5_T41 As Currency = 681279174 'Transformation Constant 41
Private Const MD5_T42 As Currency = 3936430074# 'Transformation Constant 42
Private Const MD5_T43 As Currency = 3572445317# 'Transformation Constant 43
Private Const MD5_T44 As Currency = 76029189  'Transformation Constant 44
Private Const MD5_T45 As Currency = 3654602809# 'Transformation Constant 45
Private Const MD5_T46 As Currency = 3873151461# 'Transformation Constant 46
Private Const MD5_T47 As Currency = 530742520 'Transformation Constant 47
Private Const MD5_T48 As Currency = 3299628645# 'Transformation Constant 48

'Transformation Constants - Round 4
Private Const MD5_T49 As Currency = 4096336452# 'Transformation Constant 49
Private Const MD5_T50 As Currency = 1126891415 'Transformation Constant 50
Private Const MD5_T51 As Currency = 2878612391# 'Transformation Constant 51
Private Const MD5_T52 As Currency = 4237533241# 'Transformation Constant 52
Private Const MD5_T53 As Currency = 1700485571 'Transformation Constant 53
Private Const MD5_T54 As Currency = 2399980690# 'Transformation Constant 54
Private Const MD5_T55 As Currency = 4293915773# 'Transformation Constant 55
Private Const MD5_T56 As Currency = 2240044497# 'Transformation Constant 56
Private Const MD5_T57 As Currency = 1873313359 'Transformation Constant 57
Private Const MD5_T58 As Currency = 4264355552# 'Transformation Constant 58
Private Const MD5_T59 As Currency = 2734768916# 'Transformation Constant 59
Private Const MD5_T60 As Currency = 1309151649 'Transformation Constant 60
Private Const MD5_T61 As Currency = 4149444226# 'Transformation Constant 61
Private Const MD5_T62 As Currency = 3174756917# 'Transformation Constant 62
Private Const MD5_T63 As Currency = 718787259 'Transformation Constant 63
Private Const MD5_T64 As Currency = 3951481745# 'Transformation Constant 64

Private Const BUFFER_SIZE As Integer = 1024

Private lMD5(3) As Currency
Private Count(1) As Long

'Null data used to finalise the checksum calculation
Private PADDING(63) As Byte

Private MD5Buffer As String * 64

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Initializes everything to compute MD5 hashes
'**************************************************************
    Call MD5Reset
End Sub

Public Sub MD5Reset()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Resets all private info (otherwise next MD5 will depend on the previous one)
'**************************************************************
    ' zero members
    MD5Buffer = String$(64, vbNullChar)
    Count(0) = 0
    Count(1) = 0
    
    ' Load magic state initialization constants
    lMD5(0) = MD5_INIT_STATE_0
    lMD5(1) = MD5_INIT_STATE_1
    lMD5(2) = MD5_INIT_STATE_2
    lMD5(3) = MD5_INIT_STATE_3
    
    PADDING(0) = &H80
End Sub

Public Function GetMD5File(ByVal filePath As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Computes de MD5 hash of the given file
'**************************************************************
    If dir$(filePath, vbArchive) = "" Then Exit Function
    
    Dim handle As Integer
    handle = FreeFile
    
    Open filePath For Binary Access Read Lock Write As handle

    Dim buffer(BUFFER_SIZE - 1) As Byte
    Dim i As Long
    
    'Read by blocks
    For i = 1 To LOF(handle) \ BUFFER_SIZE
        Get handle, , buffer
        
        Call UpdateChecksum(buffer, BUFFER_SIZE)
    Next i
    
    'Read reamining bytes
    If LOF(handle) Mod BUFFER_SIZE <> 0 Then
        Get handle, , buffer
        
        Call UpdateChecksum(buffer, LOF(handle) Mod BUFFER_SIZE)
    End If
    
    Close handle
    
    GetMD5File = MD5Final
End Function

Public Function GetMD5String(ByVal str As String) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Computes the MD5 hash of the given string
'**************************************************************
    'Move to byte array
    Dim arStr() As Byte
    ReDim arStr(Len(str)) As Byte
    
    'Convert string to ASCII, copy bytes to array and update
    arStr = StrConv(str, vbFromUnicode)
    Call UpdateChecksum(arStr, Len(str))
    
    GetMD5String = MD5Final
End Function

Private Function RotateLeft(ByVal num As Long, ByVal times As Byte) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Rotates left the bits of a long the given amount of times
'**************************************************************
    Dim aux As Currency
    Dim mask As Long
    
    mask = 2 ^ (32 - times) - 1
    
    'Si el primer bit es 1, (num And Not mask) es negativo, y al dividir llena a la izquierda de "F"
    'Al eliminar este bit con otro And y luego agregarlo con una suma extra se evita el problema
    If num < 0 And ((num And Not mask) Or &H80000000) Then
        aux = CCur(num And mask) * 2 ^ times + ((num And Not mask) And &H7FFFFFFF) \ 2 ^ (32 - times) + 2 ^ (times - 1)
    Else
        aux = CCur(num And mask) * 2 ^ times + (num And Not mask) \ 2 ^ (32 - times)
    End If
    
    RotateLeft = CurToLong(aux)
End Function

Private Function ShiftLeft(ByVal num As Long, ByVal times As Byte) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'VB can't shift, so we fake it
'**************************************************************
    Dim mask As Long
    
    mask = 2 ^ (32 - times - 1) - 1
    ShiftLeft = (num And mask) * 2 ^ times
    
    If num And (mask + 1) Then
        ShiftLeft = ShiftLeft Or &H80000000
    End If
End Function

Private Function ShiftRight(ByVal num As Long, ByVal times As Byte) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'VB can't shift, so we fake it
'**************************************************************
    ShiftRight = num \ 2 ^ times
End Function

Private Function CurToLong(ByVal cur As Currency) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Returns a long with the same bit configuration as an
'unsigned long would have if it had the value of the passed currency
'**************************************************************
    'We take "cur Mod (MAX_ULONG + 1)",
    'therefore we keep the least significant 32 bits
    While cur > MAX_ULONG
        cur = cur - MAX_ULONG - 1
    Wend
    
    'Now we have the "unsigned long" value,
    'we make it fix in a signed long with the same bit configuration
    'equivalent to C's memcpy( &mysignedInt, &myUnsignedInt, 4)
    If cur > MAX_LONG Then _
        cur = cur + MIN_LONG - 1 - MAX_LONG
    
    CurToLong = cur
End Function

Private Function LongToCur(ByVal num As Long) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Returns a currency with the same value as an unsigned long
'would have if it had the same bit configuration of the passed long
'**************************************************************
    If num < 0 Then
        LongToCur = MAX_ULONG + 1 + num
    Else
        LongToCur = num
    End If
End Function

Private Function ModMAX_ULONG(ByVal num As Currency) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Mod won't work on Currency, so we fake it
'**************************************************************
    While num > MAX_ULONG
        num = num - MAX_ULONG - 1
    Wend
    
    ModMAX_ULONG = num
End Function

Private Sub FF(ByRef a As Currency, ByVal b As Currency, ByVal c As Currency, ByVal d As Currency, ByVal X As Long, ByVal S As Byte, ByVal T As Currency)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'FF Transformation - MD5 specific
'**************************************************************
    Dim temp As Currency
    
    temp = LongToCur((CurToLong(b) And CurToLong(c)) Or (Not CurToLong(b) And CurToLong(d)))
    temp = CurToLong(a + X + T + temp)
    temp = LongToCur(RotateLeft(temp, S)) + b
    
    a = ModMAX_ULONG(temp)
End Sub

Private Sub GG(ByRef a As Currency, ByVal b As Currency, ByVal c As Currency, ByVal d As Currency, ByVal X As Long, ByVal S As Byte, ByVal T As Currency)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'GG Transformation - MD5 specific
'**************************************************************
    Dim temp As Currency
    
    temp = LongToCur((CurToLong(b) And CurToLong(d)) Or (CurToLong(c) And Not CurToLong(d)))
    temp = CurToLong(temp + a + X + T)
    temp = LongToCur(RotateLeft(temp, S)) + b
    
    a = ModMAX_ULONG(temp)
End Sub

Private Sub HH(ByRef a As Currency, ByVal b As Currency, ByVal c As Currency, ByVal d As Currency, ByVal X As Long, ByVal S As Byte, ByVal T As Currency)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'HH Transformation - MD5 specific
'**************************************************************
    Dim temp As Currency
    
    temp = LongToCur(CurToLong(b) Xor CurToLong(c) Xor CurToLong(d))
    temp = CurToLong(temp + a + X + T)
    temp = LongToCur(RotateLeft(temp, S)) + b
    
    a = ModMAX_ULONG(temp)
End Sub

Private Sub II(ByRef a As Currency, ByVal b As Currency, ByVal c As Currency, ByVal d As Currency, ByVal X As Long, ByVal S As Byte, ByVal T As Currency)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'II Transformation - MD5 specific
'**************************************************************
    Dim temp As Currency
    
    temp = LongToCur(CurToLong(c) Xor (CurToLong(b) Or Not CurToLong(d)))
    temp = CurToLong(temp + a + X + T)
    temp = LongToCur(RotateLeft((temp), S)) + b
    
    a = ModMAX_ULONG(temp)
End Sub

Private Sub Transform(ByRef block() As Byte)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Transforms the passed block to compute a partial hash
'**************************************************************
    Dim a As Currency
    Dim b As Currency
    Dim c As Currency
    Dim d As Currency
    
    'initialise local data with current checksum
    a = lMD5(0)
    b = lMD5(1)
    c = lMD5(2)
    d = lMD5(3)
    
    'copy bytes from input 'Block' to an array of longs 'X'
    Dim X(15) As Long
    Call CopyMemory(X(0), block(0), 64)
    
    'Perform Round 1 of the transformation
    Call FF(a, b, c, d, X(0), MD5_S11, MD5_T01)
    Call FF(d, a, b, c, X(1), MD5_S12, MD5_T02)
    Call FF(c, d, a, b, X(2), MD5_S13, MD5_T03)
    Call FF(b, c, d, a, X(3), MD5_S14, MD5_T04)
    Call FF(a, b, c, d, X(4), MD5_S11, MD5_T05)
    Call FF(d, a, b, c, X(5), MD5_S12, MD5_T06)
    Call FF(c, d, a, b, X(6), MD5_S13, MD5_T07)
    Call FF(b, c, d, a, X(7), MD5_S14, MD5_T08)
    Call FF(a, b, c, d, X(8), MD5_S11, MD5_T09)
    Call FF(d, a, b, c, X(9), MD5_S12, MD5_T10)
    Call FF(c, d, a, b, X(10), MD5_S13, MD5_T11)
    Call FF(b, c, d, a, X(11), MD5_S14, MD5_T12)
    Call FF(a, b, c, d, X(12), MD5_S11, MD5_T13)
    Call FF(d, a, b, c, X(13), MD5_S12, MD5_T14)
    Call FF(c, d, a, b, X(14), MD5_S13, MD5_T15)
    Call FF(b, c, d, a, X(15), MD5_S14, MD5_T16)

    'Perform Round 2 of the transformation
    Call GG(a, b, c, d, X(1), MD5_S21, MD5_T17)
    Call GG(d, a, b, c, X(6), MD5_S22, MD5_T18)
    Call GG(c, d, a, b, X(11), MD5_S23, MD5_T19)
    Call GG(b, c, d, a, X(0), MD5_S24, MD5_T20)
    Call GG(a, b, c, d, X(5), MD5_S21, MD5_T21)
    Call GG(d, a, b, c, X(10), MD5_S22, MD5_T22)
    Call GG(c, d, a, b, X(15), MD5_S23, MD5_T23)
    Call GG(b, c, d, a, X(4), MD5_S24, MD5_T24)
    Call GG(a, b, c, d, X(9), MD5_S21, MD5_T25)
    Call GG(d, a, b, c, X(14), MD5_S22, MD5_T26)
    Call GG(c, d, a, b, X(3), MD5_S23, MD5_T27)
    Call GG(b, c, d, a, X(8), MD5_S24, MD5_T28)
    Call GG(a, b, c, d, X(13), MD5_S21, MD5_T29)
    Call GG(d, a, b, c, X(2), MD5_S22, MD5_T30)
    Call GG(c, d, a, b, X(7), MD5_S23, MD5_T31)
    Call GG(b, c, d, a, X(12), MD5_S24, MD5_T32)

    'Perform Round 3 of the transformation
    Call HH(a, b, c, d, X(5), MD5_S31, MD5_T33)
    Call HH(d, a, b, c, X(8), MD5_S32, MD5_T34)
    Call HH(c, d, a, b, X(11), MD5_S33, MD5_T35)
    Call HH(b, c, d, a, X(14), MD5_S34, MD5_T36)
    Call HH(a, b, c, d, X(1), MD5_S31, MD5_T37)
    Call HH(d, a, b, c, X(4), MD5_S32, MD5_T38)
    Call HH(c, d, a, b, X(7), MD5_S33, MD5_T39)
    Call HH(b, c, d, a, X(10), MD5_S34, MD5_T40)
    Call HH(a, b, c, d, X(13), MD5_S31, MD5_T41)
    Call HH(d, a, b, c, X(0), MD5_S32, MD5_T42)
    Call HH(c, d, a, b, X(3), MD5_S33, MD5_T43)
    Call HH(b, c, d, a, X(6), MD5_S34, MD5_T44)
    Call HH(a, b, c, d, X(9), MD5_S31, MD5_T45)
    Call HH(d, a, b, c, X(12), MD5_S32, MD5_T46)
    Call HH(c, d, a, b, X(15), MD5_S33, MD5_T47)
    Call HH(b, c, d, a, X(2), MD5_S34, MD5_T48)

    'Perform Round 4 of the transformation
    Call II(a, b, c, d, X(0), MD5_S41, MD5_T49)
    Call II(d, a, b, c, X(7), MD5_S42, MD5_T50)
    Call II(c, d, a, b, X(14), MD5_S43, MD5_T51)
    Call II(b, c, d, a, X(5), MD5_S44, MD5_T52)
    Call II(a, b, c, d, X(12), MD5_S41, MD5_T53)
    Call II(d, a, b, c, X(3), MD5_S42, MD5_T54)
    Call II(c, d, a, b, X(10), MD5_S43, MD5_T55)
    Call II(b, c, d, a, X(1), MD5_S44, MD5_T56)
    Call II(a, b, c, d, X(8), MD5_S41, MD5_T57)
    Call II(d, a, b, c, X(15), MD5_S42, MD5_T58)
    Call II(c, d, a, b, X(6), MD5_S43, MD5_T59)
    Call II(b, c, d, a, X(13), MD5_S44, MD5_T60)
    Call II(a, b, c, d, X(4), MD5_S41, MD5_T61)
    Call II(d, a, b, c, X(11), MD5_S42, MD5_T62)
    Call II(c, d, a, b, X(2), MD5_S43, MD5_T63)
    Call II(b, c, d, a, X(9), MD5_S44, MD5_T64)

    'add the transformed values to the current checksum
    lMD5(0) = ModMAX_ULONG(lMD5(0) + a)
    lMD5(1) = ModMAX_ULONG(lMD5(1) + b)
    lMD5(2) = ModMAX_ULONG(lMD5(2) + c)
    lMD5(3) = ModMAX_ULONG(lMD5(3) + d)
End Sub

Private Function MD5Final() As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Does some finishing calculations and creates the string with the hexa hash
'**************************************************************
    'Save number of bits
    Dim bits(7) As Byte
    Call CopyMemory(bits(0), Count(0), 8)
    
    'Pad out to 56 mod 64.
    Dim index As Integer
    Dim padLen As Integer
    
    index = ShiftRight(Count(0), 3) And &H3F&
    
    If index < 56 Then
        padLen = 56 - index
    Else
        padLen = 120 - index
    End If
    
    Call UpdateChecksum(PADDING, padLen)
    
    'Append length (before padding)
    Call UpdateChecksum(bits, 8)
    
    'Store final state in 'arMD5'
    Dim arMD5(15) As Byte
    Dim longMD5(3) As Long
    Dim i As Long
    
    ' Convert Currency array to long array
    For i = 0 To 3
        longMD5(i) = CurToLong(lMD5(i))
    Next i
    
    'Copy byte blocks
    CopyMemory arMD5(0), longMD5(0), 16
    
    'Convert the hexadecimal checksum to a CString
    For i = 0 To 15
        If arMD5(i) <= 15 Then
            MD5Final = MD5Final & "0" & hex$(arMD5(i))
        Else
            MD5Final = MD5Final & hex$(arMD5(i))
        End If
    Next i
End Function

Private Sub UpdateChecksum(ByRef arInput() As Byte, ByVal InputLen As Long)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 17/06/2006
'Updates the partial hash
'**************************************************************
    Dim index As Long
    
    'Compute number of bytes mod 64
    index = ShiftRight(Count(0), 3) And &H3F&
    
    'Update number of bits
    Count(0) = Count(0) + ShiftLeft(InputLen, 3)
    If Count(0) < ShiftLeft(InputLen, 3) Then _
        Count(1) = Count(1) + 1
    
    Count(1) = Count(1) + ShiftRight(InputLen, 29)
    
    'Transform as many times as possible.
    Dim i As Long
    Dim PartLen As Integer
    PartLen = 64 - index
    If InputLen >= PartLen Then
        Dim temp() As Byte
        
        'Copy string to temp byte array
        ReDim temp(63) As Byte
        temp = StrConv(MD5Buffer, vbFromUnicode)
        
        'replace bytes in the middle
        Call CopyMemory(temp(index), arInput(0), PartLen)
        'Transform and copy back the result
        Call Transform(temp)
        MD5Buffer = StrConv(temp, vbUnicode)
        
        For i = PartLen To InputLen - 63 Step 64
            'Transform each block
            Call CopyMemory(temp(0), arInput(i), 64)
            Call Transform(temp)
            Call CopyMemory(arInput(i), temp(0), 64)
        Next i
        index = 0
    Else
        i = 0
    End If
    
    'Buffer remaining input
    If InputLen - i Then
        Dim auxBuf() As Byte
        'Copy sub block to temp array, and merge it in the partial hash
        ReDim auxBuf(InputLen - i - 1) As Byte
        Call CopyMemory(auxBuf(0), arInput(i), InputLen - i)
    
        MD5Buffer = Left$(MD5Buffer, index) & StrConv(auxBuf, vbUnicode) & Right$(MD5Buffer, 64 - InputLen + i - index)
    End If
End Sub
