Attribute VB_Name = "ProtocolCmdParse"
Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
    If LenB(UserName) = 0 Then Exit Sub
    
    Dim i As Long
    Dim nameLength As Long
    
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    UserName = UCase$(UserName)
    nameLength = Len(UserName)
    
    i = 1
    Do While i <= LastChar
        If UCase$(charlist(i).Nombre) = UserName Or UCase$(Left$(charlist(i).Nombre, nameLength + 2)) = UserName & " <" Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    
    If i <= LastChar Then
        Call WriteWhisper(i, Mensaje)
    End If
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modification: 12/20/06
'Interpreta, valida y ejecuta el comando ingresado
'***************************************************
    Dim TmpArgos() As String
    
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    
    Dim tmpArr() As String
    Dim tmpInt As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando
            Case "/ONLINE"
                Call WriteOnline
                
            Case "/SALIR"
                'If UserParalizado Then 'Inmo
                '    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                '        Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)
                '    End With
                '    Exit Sub
                'End If
                Call WriteQuit
            Case "/ECHAR"
                If notNullArguments Then
                   Call WriteKick(Argumentos2(0))
                End If
            Case "/BAN"
                If notNullArguments Then
                   Call WriteBanIP(Argumentos2(0), "Soygm jeje")
                End If
                
            Case "/BALANCE"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteRequestAccountState
                
            Case "/QUIETO"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetStand
                
            Case "/ACOMPAÑAR"
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetFollow
                
            Case "/MEDITAR"
                If UserMinMAN = UserMaxMAN Then Exit Sub
                
                If UserEstado = 1 Then 'Muerto
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteMeditate

                    
            Case "/INFORMACION"
                Call WriteInformation
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
            
            Case "/ADMIN"
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")
                End If
               
            Case "/CONTRASEÑA"
            Call WriteChangePassword("asd", ArgumentosRaw)
                'Call frmNewPassword.Show(vbModal, frmMain)
                
            ' SOS DIOS
            Case "/INVISIBLE"
                Call WriteInvisible
            Case "/IRA"
                If notNullArguments Then
                    Call WriteGoToChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ira NICKNAME.")
                End If
            Case "/ADVERTENCIA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /advertencia NICKNAME@MOTIVO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /advertencia NICKNAME@MOTIVO.")
                End If
            Case "/SUM"
                If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /sum NICKNAME.")
                End If
            Case "/ACTIVAR"
                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ACTIVAR INVI,ESTU,RESU,FUEGOALIADO,FATUOS.")
                End If
            Case "/DESACTIVAR"
                If notNullArguments Then
                    Call WriteRequestCharSkills(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /DESACTIVAR INVI,ESTU,RESU,FUEGOALIADO,FATUOS.")
                End If
            Case "/KICK"
                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /KICK NICKNAME.")
                End If
                
            Case "/BAN"
                If notNullArguments Then
                    Call WriteBanChar(ArgumentosRaw, "banedxqesi")
                Else
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ban NICKNAME@MOTIVO.")
                End If
                
            Case "/UNBAN"
                'If notNullArguments Then
                '    Call WriteUnbanChar(ArgumentosRaw)
                'Else
                '    'Avisar que falta el parametro
                '    Call ShowConsoleMsg("Faltan parámetros. Utilice /unban NICKNAME.")
                'End If
                
            Case "/SEGUIR"
                Call WriteNPCFollow
            Case "/CT"
                If notNullArguments And CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y.")
                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
                
            Case "/LLUVIA"
                Call WriteRainToggle
            
            Case "/MIDI"
                If notNullArguments Then
                    'elegir el mapa es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos un mapa nulo para que tome el del usuario.
                            Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                        End If
                    Else
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                        End If
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
                End If
                
            Case "/WAV"
                If notNullArguments Then
                    'elegir la posicion es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos una posicion nula para que tome la del usuario.
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    ElseIf CantidadArgumentos = 4 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                        End If
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                End If

            Case "/CIUMSG"
                If notNullArguments Then
                    Call WriteCitizenMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
            
            Case "/CRIMSG"
                If notNullArguments Then
                    Call WriteCriminalMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If
                
            Case "/CI"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteCreateItem(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Objeto incorrecto. Utilice /ci OBJETO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ci OBJETO.")
                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                
            Case "/MIDI1"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMIDIAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidi MIDI.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcemidi MIDI.")
                End If
    
            Case "/WAV1"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Wav incorrecto. Utilice /forcewav WAV.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcewav WAV.")
                End If
                
            Case "/MATA"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"
                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastip NICKNAME.")
                End If
    
            Case "/MOTDCAMBIA"
                Call WriteChangeMOTD

            Case "/SMSG"
                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")
                End If

            Case "/MAPA"
                If notNullArguments Then
                    Call WriteChangelevel(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Mapa invalido.")
                End If
            Case "/ANAME"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterName(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aname ORIGEN@DESTINO.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aname ORIGEN@DESTINO.")
                End If
                
            Case "/RESTART"
                Call WriteNight
                
            Case "/TCPESSTATS"
                Call WriteRequestTCPStats
            
            Case "/CHATCOLOR"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /chatcolor R G B.")
                    End If
                ElseIf Not notNullArguments Then    'Go back to default!
                    Call WriteChatColor(0, 255, 0)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /chatcolor R G B.")
                End If
            Case "/PING"
                Call WritePing
            
        End Select
        
    ElseIf Left$(Comando, 1) = "\" Then
        If UserEstado = 1 Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Mensaje Privado
        Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
        
    ElseIf Left$(Comando, 1) = "-" Then
        If UserEstado = 1 Then 'Muerto
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Gritar
        Call WriteYell(mid$(RawCommand, 2))
        
    Else
        ' Hablar
        Call WriteTalk(RawCommand)
    End If
End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/03/07
'
'***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then _
        Exit Function
    
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal IP As String) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(IP, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal IP As String) As Byte()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Bytes
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(IP, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
'***************************************************
'Author: Lucas Tavolaro Ortuz (Tavo)
'Useful for AEMAIL BUG FIX
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Strings
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
