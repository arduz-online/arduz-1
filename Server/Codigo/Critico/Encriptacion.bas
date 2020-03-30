Attribute VB_Name = "Encriptacion"
'**************************************************************
' Encriptacion.bas
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

Option Explicit

Public ELPEDIDOR As Integer 'userindex q pide proccccc

Public StringValidacion As String

Private Enum SentidoRotacion
    ROTIzquierda = 0
    ROTDerecha = 1
End Enum

Public Function ArmarStringValidacion() As String
    Dim Activo As Byte
    Dim GI As String
    Dim CI As String
    Dim PI As String
    Dim i As Byte
    Dim cascos As String
    Dim j As Byte
    Dim Key As String
    Dim MD5 As New clsMD5
    
    Key = hex$(RandomNumber(1, 65535))
    While Len(Key) < 4
        Key = "0" & Key
    Wend
    Activo = val(Trim$(GetVar(IniPath & "Server.ini", "MD5Hush", "VerArchivosCriticos")))
    If Activo = 0 Then
        ArmarStringValidacion = "0"
        Exit Function
    End If


    GI = UCase$(MD5.GetMD5File(App.Path & "\criticos\graficos.ind"))
    Call MD5.MD5Reset
    CI = UCase$(MD5.GetMD5File(App.Path & "\criticos\cabezas.ind"))
    Call MD5.MD5Reset
    PI = UCase$(MD5.GetMD5File(App.Path & "\criticos\personajes.ind"))
    Call MD5.MD5Reset
    cascos = UCase$(MD5.GetMD5File(App.Path & "\criticos\cascos.ind"))

    If InStr(1, GI & CI & PI & cascos, "FILE NOT FOUND") > 0 Then
        Call MsgBox("No se encuentran los archivos graficos.ind, cabezas.ind, personajes.ind, cascos.ind en la carpeta .\criticos", vbCritical, "Error de validación")
        ArmarStringValidacion = "0"
        Exit Function
    End If

    For i = 1 To Len(GI)
        ArmarStringValidacion = ArmarStringValidacion & hex$(hexHex2Dec(mid$(GI, i, 1)) Xor hexHex2Dec(mid$(CI, i, 1)) Xor hexHex2Dec(mid$(PI, i, 1)) Xor hexHex2Dec(mid$(cascos, i, 1)) Xor hexHex2Dec(mid$(Key, (i Mod 4) + 1, 1)))
    Next i
    
    ArmarStringValidacion = Key & ArmarStringValidacion 'rota con cada WS
    
    Set MD5 = Nothing
End Function

Public Function ProtoCrypt(ByVal S As String, ByVal Key As Integer) As String
'esta funcion encripta un string

On Error GoTo errorHandlerEncriptar
Dim L As Integer
Dim tr As Byte
Dim rap As Integer
Dim i As Long
Dim ascii As Integer
Dim enc As Long
Dim B1 As Byte
Dim B2 As Byte
Dim Encriptar As String

Key = UserList(Key).KeyCrypt And &HFF

L = Len(S)
tr = RandomNumber(0, 1)
rap = RandomNumber(1, L)

If rap And &H1 Then rap = rap + 1
If L <= &HFF Then
    If tr = 1 Then
        'de atras pa delante (LCBBBB..BBSSSSS)
        Encriptar = hexMd52Asc("02" & Right$("0" & hex$(L), 2))
        For i = 1 To rap
            Encriptar = Encriptar & hexMd52Asc(hex$(RandomNumber(56, 255)))
        Next i
        For i = 1 To L
            ascii = Asc(mid$(S, i, 1))
            'enc = ascii * 16
            B1 = ascii And &HF
            B2 = ascii And &HF0
            B2 = B2 \ 16
            enc = (B1 Xor B2)
            'enc = enc Or ((B1 And B2) * 4096) Or &H4006 'el ultimo or es para no obtener ENDC en el dato!
            enc = enc Or ((B1 And B2) * 4096)
            'lista la basura, ahora el char
            ascii = RotateCarry(ascii, i And &H7, SentidoRotacion.ROTDerecha)
            enc = enc Or ((ascii Xor Key) * 16)
            If enc < &H1FF Then
                Encriptar = Encriptar & hexMd52Asc("40" & hex$(enc))
            Else
                Encriptar = Encriptar & hexMd52Asc(hex$(enc))
            End If
        Next i
        ProtoCrypt = Encriptar
    Else
        'de delante pa atras (LCSSSSSBBBB..BB)
        Encriptar = hexMd52Asc("04" & Right$("0" & hex$(L), 2))
        
        For i = L To 1 Step -1
            ascii = Asc(mid$(S, i, 1))
            'enc = ascii * 16
            B1 = ascii And &HF
            B2 = ascii And &HF0
            B2 = B2 \ 16
            enc = (B1 Xor B2)
            enc = enc Or ((B1 And B2) * 4096) Or &H4006
            'lista la basura, ahora el char
            ascii = RotateCarry(ascii, i And &H7, ROTDerecha)
            enc = enc Or ((ascii Xor Key) * 16)
            If enc < &H1FF Then
                Encriptar = Encriptar & hexMd52Asc("30" & hex$(enc))
            Else
                Encriptar = Encriptar & hexMd52Asc(hex$(enc))
            End If
        Next i
        For i = 1 To rap
            Encriptar = Encriptar & hexMd52Asc(hex$(RandomNumber(45, 255)))
        Next i
        ProtoCrypt = Encriptar
    End If
Else
    'esto es acsolutamente imposible que suceda
    ProtoCrypt = S
End If

Exit Function
errorHandlerEncriptar:
'Call LogError("Error encriptando: " & s)
ProtoCrypt = S

End Function

Private Function RotateCarry(ByVal b As Byte, ByVal veces As Byte, ByVal Sentido As SentidoRotacion) As Byte
Dim resto As Byte

Do While veces > 0
    If Sentido = ROTDerecha Then
        resto = b And &H1
        b = b \ 2
        If resto = 1 Then
            b = b Or &H80
        End If
    Else
        If (b And &H80) > 0 Then
            b = (b And &H7F) * 2 + 1
        Else
            b = b * 2
        End If
    End If
    veces = veces - 1
Loop

RotateCarry = b

End Function

Public Function jeringoso(ByVal L As String) As String
Dim i As Integer
Dim unaletra As String
Dim buenaLetra As String

For i = 1 To Len(L)
    unaletra = mid$(L, i, 1)

    Select Case unaletra
        Case "%"
        buenaLetra = "A"
        Case "#"
        buenaLetra = "B"
        Case "&"
        buenaLetra = "C"
        Case "("
        buenaLetra = "D"
        Case "{"
        buenaLetra = "E"
        Case "}"
        buenaLetra = "F"
        Case "!"
        buenaLetra = "G"
        
        Case ">"
        buenaLetra = "H"
        
        Case "."
        buenaLetra = "I"
        
        Case "g"
        buenaLetra = "J"
        
        Case "d"
        buenaLetra = "K"
        
        Case "<"
        buenaLetra = "L"
        
        Case "s"
        buenaLetra = "M"
        
        Case "f"
        buenaLetra = "N"
        
        Case "F"
        buenaLetra = "O"
        
        Case "r"
        buenaLetra = "P"
        
        Case "P"
        buenaLetra = "Q"
        
        Case "K"
        buenaLetra = "R"
        
        Case "5"
        buenaLetra = "S"
        
        Case "8"
        buenaLetra = "T"
        
        Case "2"
        buenaLetra = "U"
        
        Case "0"
        buenaLetra = "V"
        
        Case "_"
        buenaLetra = "W"
        
        Case "="
        buenaLetra = "X"
        
        Case "x"
        buenaLetra = "Y"
        
        Case "m"
        buenaLetra = "Z"
        Case "Q"
        buenaLetra = "1"
        Case "W"
        buenaLetra = "2"
        Case "E"
        buenaLetra = "3"
        Case "R"
        buenaLetra = "4"
        Case "T"
        buenaLetra = "5"
        Case "Y"
        buenaLetra = "6"
        Case "U"
        buenaLetra = "7"
        Case "O"
        buenaLetra = "8"
        Case "A"
        buenaLetra = "9"
        Case "S"
        buenaLetra = "0"
        Case "["
        buenaLetra = "-"
        Case "]"
        buenaLetra = ":"
        Case "/"
        buenaLetra = "."
      
        
        Case Else
        buenaLetra = "?"

    End Select

    jeringoso = jeringoso & buenaLetra
Next i


End Function

Public Function MoveCharCrypt(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
Dim b(1 To 4) As Byte
Dim Kk  As Integer
Dim ChI As Integer
    'PARAMETROS:
    'Userindex: Receptor del mensaje MP
    'Charindex: Personaje que se mueve
    'X,Y: Nuevas coordenadas


    'para el char necesito 13 bits para cubrir los 10k chars posibles
    'para las coordenadas, puedo sacrificar un bit por coordenada
    'entonces la codificacion la hacemos, para evitar los ENDC de la siguiente manera:
    
    'coord x= X8X7X6X5X4X3X2X1
    'coord y= Y8Y7Y6Y5Y4Y3Y2Y1
    'charind= C16C15C14.....C1
    
    '   1   C3  C2  C1  X4  X3  X2  X1
    '   X7  X6  C7  C6  X5  1   C5  C4
    '   C10 C9  1   C8  Y4  Y3  Y2  Y1
    '   C14 C12 C13 C11 1   Y7  Y6  Y5
    
    'EL 1 mezclado en los bits, confunde y evita los endC
    'ANTES, a coordx,coordy y charind se le hace xor con la key privada de cada user
    
    Kk = UserList(UserIndex).KeyCrypt
    X = X Xor Kk
    Y = Y Xor Kk
    ChI = CharIndex Xor Kk


    b(1) = (X And &HF)
    b(2) = (X And &H10) / 2 + (X And &H60) * 2
    b(3) = (Y And &HF)
    b(4) = ((Y And &H70) / 16)
    'x x + +  + + + +   + + + +  + + + +
    b(1) = b(1) + ((ChI And &H7) * 16) + &H80                           '80 = bit anti ENDc
    b(2) = b(2) + ((ChI And &H18) / 8) + ((ChI And &H60) / 2) + &H4     '04 = bit anti ENDc
    b(3) = b(3) + ((ChI And &H80) / 8) + ((ChI And &H300) / 4) + &H20   '20 = bit anti ENDc
    b(4) = b(4) + (ChI And &H3C00) / 64 + &H8

    MoveCharCrypt = Chr$(b(1)) & Chr$(b(2)) & Chr$(b(3)) & Chr$(b(4))

End Function

Public Function MoveNPCCrypt(ByVal NpcIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
Dim b(1 To 4) As Byte
Dim Kk  As Integer
Dim ChI As Integer

    'C = charindex
    'byte1  CCCC 1XXX    Char byte1 parte 1, X byte 1
    'byte2  1CCC XXXX    Char byte1 parte 2, X byte 2
    'byte3  Y1YY CCCC     Char byte2 parte 1, Y byte 1
    'byte4  C1CC YYYY    Char byte2 parte 2, Y byte 2
    'nunca puede formarse un byte ENDC = 1
    'Debug.Print "Codificando: " & Npclist(NpcIndex).Char.charindex & " " & X & "," & Y
    Kk = 17320 Or &H1248    'para evitar endc's
    X = X Xor (Kk And &HFF)
    Y = Y Xor (Kk And &HFF)
    ChI = Npclist(NpcIndex).Char.CharIndex Xor Kk

    b(1) = ((X And &H70) / 16) Or 8
    b(2) = (X And &HF)
    b(3) = (Y And &H30) + (Y And &H40) * 2 + &H40
    b(4) = (Y And &HF)
    b(1) = b(1) + (ChI And &H3C00) / 64
    b(2) = b(2) + (ChI And &H380) / 8 + &H80
    b(3) = b(3) + (ChI And &H78) / 8
    b(4) = b(4) + ((ChI And &H3) * 16) + (ChI And &H4) * 32 + &H40  '40 para endc's
    MoveNPCCrypt = Chr$(b(1)) & Chr$(b(2)) & Chr$(b(3)) & Chr$(b(4))
    'xx11 1122 2333 3444
    'xxxx 1xxx  1xxx xxxx  1xxx xxxx  x1xx xxxx
End Function

