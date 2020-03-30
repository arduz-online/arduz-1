Attribute VB_Name = "ProtocolEx"
'**************************************************************
' ProtocolEx.bas - Handles expanded messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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
'Expands Protocol.bas adding all security related (and therefore private)
'packets' handling.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

Private Const EXTENDED_PROTOCOL_PACKET_ID As Byte = 246

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue

''
'Determines the current mode of the anticheats.
Private ModoAntiCheats As Byte

''
'Determines wether the anticheats is online or not.
Private AntiCheatsOnline As Boolean

''
'MD5 expected to be received by the anticheats.
Private AntiCheatsEsperado As String

' ATTENTION : First packet here MUST have a value of the LAST packet in ServerPacketID enum + 1!!
Private Enum ServerPacketIDEx
    IsProcessRunning = 106  ' WIUW
    RequestPressedKeysLog   ' DLDT
    RequestProcessList      ' DLDP
    RequestMD5              ' RMDC
    CRCCheck
    ValidateClient          ' VAL
End Enum

' ATTENTION : These headers are 2 bytes long, first one is EXTENDED_PROTOCOL_PACKET_ID (first index after last ClientPacketID), second one indicates the secutiry packet.
Private Enum ClientPacketIDEx
    debuggerDetected = 1            ' DD
    SpeedHackDetected               ' SHD
    MD5Reported                     ' RMDC
    ReportProcessList               ' DLDP
    ReportKeyLog                    ' DLDT
    ProcessIsRunning                ' IAUT
    CRCCheckSum
    
    'GM Commands
    GlobalAnticheats                '/GLOBALANTICHEATS
    SetAnticheats                   '/SETANTICHEATS
    TurnOnAnticheats                '/PRENDERANTICHEATS
    RequestMD5                      '/MD5
    RequestProcessList              '/PRO
    WhoIsUsing                      '/QUIENUSA
    RequestKeyLog                   '/CHI
    SeeReportedMD5                  '/SMD5
    AnticheatsBan                   '/ANTICHEATS
End Enum

''
' Handles incoming data extended commands (Alkon's security).
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingDataEx(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
    'Packet IDs are 2 bytes long here, so we ignore them
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Obtain the second byte in the buffer (VB uses little-endian)
    Select Case UserList(UserIndex).incomingData.PeekInteger() \ &H100
        Case ClientPacketIDEx.debuggerDetected       ' DD
            Call HandleDebuggerDetected(UserIndex)
        
        Case ClientPacketIDEx.SpeedHackDetected      ' SHD
            Call HandleSpeedHackDetected(UserIndex)
        
        Case ClientPacketIDEx.MD5Reported            ' RMDC
            Call HandleMD5Reported(UserIndex)
        
        Case ClientPacketIDEx.ReportProcessList      ' DLDP
            Call HandleReportProcessList(UserIndex)
        
        Case ClientPacketIDEx.ReportKeyLog           ' DLDT
            Call HandleReportKeyLog(UserIndex)
        
        Case ClientPacketIDEx.ProcessIsRunning       ' IAUT
            Call HandleProcessIsRunning(UserIndex)
        
        Case ClientPacketIDEx.CRCCheckSum
            Call HandleCRCCheckSum(UserIndex)
        
        Case ClientPacketIDEx.GlobalAnticheats        '/GLOBALANTICHEATS
            Call HandleGlobalAnticheats(UserIndex)
        
        Case ClientPacketIDEx.SetAnticheats           '/SETANTICHEATS
            Call HandleSetAnticheats(UserIndex)
        
        Case ClientPacketIDEx.TurnOnAnticheats        '/PRENDERANTICHEATS
            Call HandleTurnOnAnticheats(UserIndex)
        
        Case ClientPacketIDEx.RequestMD5              '/MD5
            Call HandleRequestMD5(UserIndex)
        
        Case ClientPacketIDEx.RequestProcessList      '/PRO
            Call HandleRequestProcessList(UserIndex)
        
        Case ClientPacketIDEx.WhoIsUsing              '/QUIENUSA
            Call HandleWhoIsUsing(UserIndex)
        
        Case ClientPacketIDEx.RequestKeyLog           '/CHI
            Call HandleRequestKeyLog(UserIndex)
        
        Case ClientPacketIDEx.SeeReportedMD5          '/SMD5
            Call HandleSeeReportedMD5(UserIndex)
        
        Case ClientPacketIDEx.AnticheatsBan           '/ANTICHEATS
            Call HandleAnticheatsBan(UserIndex)
        
        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
    End Select
End Sub

''
' Handles the "DebuggerDetected" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDebuggerDetected(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim CharFile As String
    Dim Count As Integer
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    'Se detectó un debugger!!
    Select Case ModoAntiCheats
        Case 0
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó anticheats", FontTypeNames.FONTTYPE_SERVER))
        Case 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó el control anticheats.", FontTypeNames.FONTTYPE_SERVER))
        Case 2
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats echó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_SERVER))
            Call CloseSocket(UserIndex)
        Case 3
            Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "DEBUGGER PRESENTE")
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats banneó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_FIGHT))
            UserList(UserIndex).flags.Ban = 1
            CharFile = CharPath & UserList(UserIndex).name & ".chr"
            Call WriteVar(CharFile, "FLAGS", "Ban", "1")
            'ponemos la pena
            Count = val(GetVar(CharFile, "PENAS", "Cant"))
            Call WriteVar(CharFile, "PENAS", "Cant", Count + 1)
            Call WriteVar(CharFile, "PENAS", "P" & Count + 1, "Sistema Anti Cheats: BAN POR DEBUGGER " & Date & " " & time)
            Call CloseSocket(UserIndex)
    End Select
End Sub

''
' Handles the "SpeedHackDetected" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpeedHackDetected(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim CharFile As String
    Dim Count As Integer
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    'Se detectó un speedhack!!
    Select Case ModoAntiCheats
        Case 0
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó anticheats", FontTypeNames.FONTTYPE_SERVER))
        Case 1
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó el control anticheats.", FontTypeNames.FONTTYPE_SERVER))
        Case 2
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats echó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_SERVER))
            Call CloseSocket(UserIndex)
        Case 3
            Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "SPEEDHACK")
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats banneó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_FIGHT))
            UserList(UserIndex).flags.Ban = 1
            CharFile = CharPath & UserList(UserIndex).name & ".chr"
            Call WriteVar(CharFile, "FLAGS", "Ban", "1")
            'ponemos la pena
            Count = val(GetVar(CharFile, "PENAS", "Cant"))
            Call WriteVar(CharFile, "PENAS", "Cant", Count + 1)
            Call WriteVar(CharFile, "PENAS", "P" & Count + 1, "Sistema Anti Cheats: BAN POR SPEEDHACK " & Date & " " & time)
            Call CloseSocket(UserIndex)
    End Select
End Sub

''
' Handles the "MD5Reported" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMD5Reported(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    Dim CharFile As String
    Dim Count As Integer
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    UserList(UserIndex).flags.MD5Reportado = buffer.ReadASCIIString()
    
    If AntiCheatsOnline Then
        'Keep only the MD5, discard remote ip and port
        UserList(UserIndex).flags.MD5Reportado = Left$(UserList(UserIndex).flags.MD5Reportado, 32)
        
        If UserList(UserIndex).flags.MD5Reportado <> AntiCheatsEsperado Then
            Select Case ModoAntiCheats
                Case 0
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó anticheats", FontTypeNames.FONTTYPE_SERVER))
                Case 1
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UCase$(UserList(UserIndex).name) & " reprobó el control anticheats.", FontTypeNames.FONTTYPE_SERVER))
                Case 2
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats echó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                Case 3
                    Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "CLIENTE EDITADO")
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats banneó a " & UCase$(UserList(UserIndex).name), FontTypeNames.FONTTYPE_FIGHT))
                    UserList(UserIndex).flags.Ban = 1
                    CharFile = CharPath & UserList(UserIndex).name & ".chr"
                    Call WriteVar(CharFile, "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharFile, "PENAS", "Cant"))
                    Call WriteVar(CharFile, "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharFile, "PENAS", "P" & Count + 1, "Sistema Anti Cheats: BAN POR CLIENTE EDITADO " & Date & " " & time)
                    Call CloseSocket(UserIndex)
            End Select
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ReportProcessList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReportProcessList(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim data As String
    
    data = buffer.ReadASCIIString()
    
    Call WriteConsoleMsg(ELPEDIDOR, jeringoso(data), FontTypeNames.FONTTYPE_INFO)
    
    Call FlushBuffer(ELPEDIDOR)
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ReportKeyLog" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReportKeyLog(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim data As String
    Dim tUser As Integer
    
    data = buffer.ReadASCIIString()
    
    tUser = NameIndex("MARAXUS")
    If tUser <> 0 Then
        Call WriteConsoleMsg(tUser, ">" & data, FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(tUser)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "ProcessIsRunning" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleProcessIsRunning(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    Call WriteConsoleMsg(ELPEDIDOR, UserList(UserIndex).name & " dió positivo!", FontTypeNames.FONTTYPE_INFO)
    
    Call FlushBuffer(ELPEDIDOR)
End Sub

''
' Handles the "CRCCheckSum" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCRCCheckSum(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    'Reset tolerance to 0
    UserList(UserIndex).Security.Md5RequestedSecurityTolerance = 0
    
    If UserList(UserIndex).Security.ExpectedCheckSum <> UserList(UserIndex).incomingData.ReadLong() Then
        'FAILED! Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    Else
        'Set flag
        UserList(UserIndex).Security.CheckSumValidated = True
    End If
End Sub

''
' Handles the "GlobalAnticheats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGlobalAnticheats(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim mode As Byte
    Dim file As String
    
    mode = buffer.ReadByte()
    file = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        ModoAntiCheats = mode
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRequestMD5(file))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SetAnticheats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetAnticheats(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 34 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    Dim expected As String
    
    expected = UserList(UserIndex).incomingData.ReadASCIIStringFixed(32)
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        AntiCheatsEsperado = UCase$(expected)
        
        Call WriteConsoleMsg(UserIndex, "String expected: " & expected, FontTypeNames.FONTTYPE_SERVER)
    End If
End Sub

''
' Handles the "TurnOnAnticheats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTurnOnAnticheats(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger
    
    Dim mode As Byte
    
    mode = UserList(UserIndex).incomingData.ReadByte
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        Select Case mode
            Case 0
                AntiCheatsOnline = False
            
            Case 1
                AntiCheatsOnline = False
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats ha sido desconectado (por ahora)", FontTypeNames.FONTTYPE_SERVER))
            
            Case 2
                AntiCheatsOnline = True
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anticheats ha sido activado", FontTypeNames.FONTTYPE_SERVER))
            
            Case 3
                AntiCheatsOnline = True
        End Select
    End If
End Sub

''
' Handles the "RequestMD5" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMD5(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim UserName As String
    Dim filename As String
    Dim tUser As Integer
    
    UserName = buffer.ReadASCIIString()
    filename = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        tUser = NameIndex(UserName)
        If tUser = 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario no conectado", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteRequestMD5(tUser, filename)
            Call FlushBuffer(tUser)
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestProcessList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestProcessList(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim UserName As String
    Dim tUser As Integer
    
    UserName = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        tUser = NameIndex(UserName)
        If tUser = 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario no conectado", FontTypeNames.FONTTYPE_INFO)
        Else
            ELPEDIDOR = UserIndex
            Call WriteRequestProcessList(tUser)
            Call FlushBuffer(tUser)
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "WhoIsUsing" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhoIsUsing(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim filename As String
    Dim tUser As Integer
    
    filename = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        ELPEDIDOR = UserIndex
        Call WriteConsoleMsg(UserIndex, "NO DESCONECTARSE DURANTE 1 MINUTO DESPUES DE USAR ESTE COMANDO!", FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageIsProcessRunning(filename))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "RequestKeyLog" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestKeyLog(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim UserName As String
    Dim tUser As Integer
    
    UserName = buffer.ReadASCIIString()
    
    If UCase$(UserList(UserIndex).name) = "MARAXUS" Then
        tUser = NameIndex(UserName)
        
        If tUser = 0 Then
            Call WriteConsoleMsg(UserIndex, "Usuario no conectado", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteRequestPressedKeysLog(tUser)
            Call FlushBuffer(tUser)
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "SeeReportedMD5" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSeeReportedMD5(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim UserName As String
    Dim tUser As Integer
    
    UserName = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        tUser = NameIndex(UserName)
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario está offline.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "El MD5 pedido es: " & UserList(tUser).flags.MD5Reportado, FontTypeNames.FONTTYPE_INFO)
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Handles the "AnticheatsBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAnticheatsBan(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadInteger
    
    Dim UserName As String
    Dim reason As String
    Dim tUser As Integer
    Dim priv As Long
    Dim Count As Integer
    Dim rank As Integer
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    UserName = buffer.ReadASCIIString()
    reason = buffer.ReadASCIIString()
    
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.RoleMaster) <> 0 And (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
        tUser = NameIndex(UserName)
        
        If tUser <= 0 Then
            Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_TALK)
            
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                priv = UserDarPrivilegioLevel(UserName)
                
                If (priv And rank) > (UserList(UserIndex).flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No podés banear a alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                ElseIf GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                    Call WriteConsoleMsg(UserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call Ban(UCase$(UserName), "Sistema Anti Cheats", UCase$(reason))
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anti cheats baneó a " & UCase$(UserName) & "(" & UCase$(reason) & ")", FontTypeNames.FONTTYPE_FIGHT))
                    
                    'ponemos el flag de ban a 1
                    Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                    'ponemos la pena
                    Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, "Sistema anti cheats: BAN POR " & LCase$(reason) & " " & Date & " " & time)
                    
                    If priv > 0 Then
                        UserList(UserIndex).flags.Ban = 1
                        Call CloseSocket(UserIndex)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    End If
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And rank) > (UserList(UserIndex).flags.Privilegios And rank) Then
                Call WriteConsoleMsg(UserIndex, "No podés banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call Ban(UCase$(UserName), "Sistema Anti Cheats", UCase$(reason))
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> El sistema anti cheats baneó a " & UCase$(UserName) & "(" & UCase$(reason) & ")", FontTypeNames.FONTTYPE_FIGHT))
                
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
                
                If UserList(tUser).flags.Privilegios > PlayerType.User Then
                    UserList(UserIndex).flags.Ban = 1
                    Call CloseSocket(UserIndex)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                End If
                
                Call LogGM(UserList(UserIndex).name, "BAN a " & UserList(tUser).name)
                
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, "Sistema Anti Cheats: BAN POR " & LCase$(reason) & " " & Date & " " & time)
                
                Call CloseSocket(tUser)
            End If
        End If
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:
    Dim error As Long
    error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If error <> 0 Then _
        Err.raise error
End Sub

''
' Writes the "IsProcessRunning" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    processName Process we are looking for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Sub WriteIsProcessRunning(ByVal UserIndex As Integer, ByVal processName As String)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IsProcessRunning" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageIsProcessRunning(processName))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RequestPressedKeysLog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Sub WriteRequestPressedKeysLog(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPressedKeysLog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketIDEx.RequestPressedKeysLog)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RequestProcessList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Sub WriteRequestProcessList(ByVal UserIndex As Integer)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestProcessList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketIDEx.RequestProcessList)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RequestMD5" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    fileName The file whose MD5 is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Sub WriteRequestMD5(ByVal UserIndex As Integer, ByVal filename As String)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMD5" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    'A MD5 was requested, set tolerance to one minute for CRC checksums
    UserList(UserIndex).Security.Md5RequestedSecurityTolerance = MD5_REQUEST_SECURITY_TOLERANCE
    
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRequestMD5(filename))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CRCCheck" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    algorithm   The CheckSum algorithm to be used.
' @param    key         The random key to use.
' @param    data        The random data whose checksum will be computed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCRCCheck(ByVal UserIndex As Integer, ByVal algorithm As CheckSumType, ByVal key As Long, ByVal data As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CRCCheck" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketIDEx.CRCCheck)
        
        Call .WriteByte(algorithm)
        Call .WriteLong(key)
        Call .WriteLong(data)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ValidateClient" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteValidateClient(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ValidateClient" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketIDEx.ValidateClient)
        
        Call .WriteASCIIStringFixed(Encriptacion.StringValidacion)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Prepares the "RequestMD5" message to the given user's outgoing data buffer.
'
' @param    fileName The file whose MD5 is requested.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Function PrepareMessageRequestMD5(ByVal filename As String) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RequestMD5" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketIDEx.RequestMD5)
        Call .WriteASCIIString(filename)
        
        PrepareMessageRequestMD5 = .ReadASCIIStringFixed(.length)
    End With
End Function

''
' Prepares the "IsProcessRunning" message and returns it.
'
' @param    processName Process we are looking for.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Private Function PrepareMessageIsProcessRunning(ByVal processName As String) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "IsProcessRunning" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketIDEx.IsProcessRunning)
        Call .WriteASCIIString(processName)
        
        PrepareMessageIsProcessRunning = .ReadASCIIStringFixed(.length)
    End With
End Function
