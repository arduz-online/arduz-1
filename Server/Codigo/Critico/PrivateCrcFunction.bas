Attribute VB_Name = "PrivateCrcFunction"
'**************************************************************
' PrivateCrcFunction.bas - Generates CRC checksums. This module
' compromises security. It should NEVER be distributed.
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
' Has different CRC algorithms.
'  - Original CRC32     - (Unknown)
'  - VB.Net CRC32       - Paul (wpsjr1@succeed.net) - It was modified to receive strings instead of byte arrays by Maraxus
'  - Adler32            - Maraxus (dso_maraxus@hotmail.com) - Translated from C++, the original code is used by ZLib.
'  - CRC16              - Salvito
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

Public Enum CheckSumType
    CST_Adler32 = 0
    CST_CRC32
    CST_CRC32Net
    CST_CRC16
End Enum

''
' Public key. Used by CRC32.
Public MixedKey As Long

''
' Adler32 related constants
Private Const BASE As Long = 65521      ' largest prime smaller than 65536
Private Const NMAX As Integer = 5552    ' NMAX is the largest n such that 255n(n+1)/2 + (n+1)(BASE-1) <= 2^32-1

''
' CRC32Net related variables
Private crc32Table() As Long
Private crc32NetInitialized As Boolean

Private Function Adler32(ByVal adler As Long, ByVal buf As String, ByVal length As Integer) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 3/06/2006
'Computes an adler32 checksum
'**************************************************************
    Dim sum2 As Long
    Dim N As Long
    Dim Start As Long
    Dim i As Long
    
    'Divido el adler en 2 partes de 16 bits
    sum2 = adler \ &H10000
    adler = adler And &HFFFF&
    
    If buf = vbNullString Then
        Adler32 = 1
        Exit Function
    End If
    
    Start = 1
    
    ' do length NMAX blocks -- requires just one modulo operation
    While length >= NMAX
        length = length - NMAX
        N = NMAX \ 16 - 1           ' NMAX is divisible by 16 */
        Do While N
            For i = 0 To 15
                adler = adler + Asc(mid$(buf, Start + i, 1))
                sum2 = sum2 + adler
            Next i
            Start = Start + 16
            N = N - 1
        Loop
        
        adler = adler Mod BASE
        sum2 = sum2 Mod BASE
    Wend

    ' do remaining bytes (less than NMAX, still just one modulo)
    If length Then  ' avoid modulos if none remaining
        While length >= 16
            length = length - 16
            For i = 0 To 15
                adler = adler + Asc(mid$(buf, Start + i, 1))
                sum2 = sum2 + adler
            Next i
            Start = Start + 16
        Wend
        
        While length
            length = length - 1
            adler = adler + Asc(mid$(buf, Start, 1))
            Start = Start + 1
            sum2 = sum2 + adler
        Wend
        
        adler = adler Mod BASE
        sum2 = sum2 Mod BASE
    End If

    ' return recombined sums
    If sum2 Or &H8000& Then
        'Avoid overflows if first bit is set
        Adler32 = ((sum2 And &H7FFF&) * &H10000 + adler) Or &H80000000
    Else
        Adler32 = sum2 * &H10000 + adler
    End If
End Function

Private Function GenCrC(ByVal key As Long, ByVal sdData As String) As Long
'**************************************************************
'Author: Unkown
'Last Modify Date: ??/??/????
'Computes a CRC32 checksum
'**************************************************************
    Dim CrcKey As Long
    Dim MiniCrcKey As Byte
    Dim CrcKey1 As Long
    Dim MiniCrcKey1 As Byte
    Dim Letra As Byte
    Dim cadena As String
    Dim Cadena1 As String
    Dim Contador As Long
    Dim i As Integer
    Dim signBits As Long
    
    CrcKey = key
    cadena = sdData
    Cadena1 = cadena
    CrcKey1 = MixedKey
    
    Do While Len(Cadena1) > 0
        Letra = Asc(Left$(Cadena1, 1))
        Letra = Not Letra Xor (CrcKey And 200)
        Cadena1 = Right$(Cadena1, (Len(Cadena1) - 1))
        MiniCrcKey1 = CrcKey1 And 248
        CrcKey1 = CrcKey1 And &HFFAFFF00
        MiniCrcKey1 = MiniCrcKey1 And Letra
        CrcKey1 = CrcKey1 Or MiniCrcKey1
        MiniCrcKey1 = MiniCrcKey1 Mod (32)
        
        If MiniCrcKey1 <> 0 Then
            For i = 1 To MiniCrcKey1
                signBits = CrcKey1 And &H80000001
                CrcKey1 = ((CrcKey1 And &H7FFFFFFE) \ 2) Or ((signBits < 0) And &H40000000) Or (CBool(signBits And 1) And &H80000000)
            Next
        End If
        
        CrcKey1 = Not CrcKey1
        MiniCrcKey1 = CrcKey1 And 255
        
        If MiniCrcKey1 <> 0 Then
            For i = 1 To MiniCrcKey1
                signBits = CrcKey And &HC0000000
                CrcKey = ((CrcKey And &H3FFFFFFF) * 2) Or ((signBits < 0) And 1) Or (CBool(signBits And &H40000000) And &H80000000)
            Next
        End If
        
        MiniCrcKey = CrcKey And 251
        CrcKey = CrcKey - MiniCrcKey
        MiniCrcKey = MiniCrcKey Xor (Letra Or MiniCrcKey1)
        CrcKey = CrcKey Or MiniCrcKey
    Loop
    
    CrcKey = CrcKey Xor CrcKey1
    CrcKey1 = CrcKey
    CrcKey = CrcKey And 65535
    
    For i = 1 To 16
        signBits = CrcKey1 And &HC0000000
        CrcKey1 = ((CrcKey1 And &H3FFFFFFF) * 2) Or ((signBits < 0) And 1) Or (CBool(signBits And &H40000000) And &H80000000)
    Next i
    
    CrcKey = CrcKey Xor (CrcKey1 And 65535)

    GenCrC = CrcKey
End Function

Private Function CRC32Net(ByVal data As String) As Long
'**************************************************************
'Author: Paul (wpsjr1@succeed.net)
'Last Modify Date: 2/10/2007
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'Initializesthe CRC32 tables for VB.Net's CRC32 algorithm
'**************************************************************
    Dim crc32Result As Long
    Dim buffer() As Byte
    Dim i As Long
    Dim iLookup As Integer
    
    'Transform string to array
    buffer = StrConv(data, vbFromUnicode)
    
    crc32Result = &HFFFFFFFF
    
    ' Initialize CRC tables if not done already
    If Not crc32NetInitialized Then CRC32NetInitialize
    
    For i = LBound(buffer()) To UBound(buffer())
        iLookup = (crc32Result And &HFF) Xor buffer(i)
        crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215 ' nasty shr 8 with vb :/
        crc32Result = crc32Result Xor crc32Table(iLookup)
    Next i
    
    CRC32Net = Not crc32Result
End Function

Private Sub CRC32NetInitialize()
'**************************************************************
'Author: Paul (wpsjr1@succeed.net)
'Last Modify Date: ??/??/????
'Initializes the CRC32 tables for VB.Net's CRC32 algorithm
'**************************************************************
    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Long
    Dim i As Integer, j As Integer
    Dim dwCrc As Long
    
    ReDim crc32Table(256) As Long
    
    dwPolynomial = &HEDB88320
    
    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i
    
    crc32NetInitialized = True
End Sub

Private Function CRC16(ByVal key As Long, ByVal data As String) As Integer
'**************************************************************
'Author: Salvito
'Last Modify Date: 2/07/2007
'Computes a custom CRC16 designed by Alejandro Salvo
'**************************************************************
    Dim i As Long
    Dim vstr() As Byte
    Dim SumaEspecialDeCaracteres As Long
    
    vstr = StrConv(data, vbFromUnicode)
    
    For i = 0 To Len(data) - 1
        SumaEspecialDeCaracteres = SumaEspecialDeCaracteres + vstr(i) * (1 + key - i)
    Next i
    
    CRC16 = CInt(Abs(SumaEspecialDeCaracteres) And &HFFFF&)
End Function

Public Function CheckSum(ByVal algorithm As CheckSumType, ByVal key As Long, ByVal buf As String) As Long
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 7/02/2007
'Computes a checksum using the selected algorithm
'**************************************************************
    Select Case algorithm
        Case CST_Adler32
            CheckSum = Adler32(key, buf, Len(buf))
        
        Case CST_CRC32
            CheckSum = GenCrC(key, buf)
        
        Case CST_CRC32Net
            CheckSum = CRC32Net(buf)
        
        Case CST_CRC16
            CheckSum = CRC16(key, buf)
    End Select
End Function
