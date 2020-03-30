Attribute VB_Name = "Admin"

Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public ReiniciarServer As Long

Public tInicioServer As Long

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long

'BALANCE

Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Public Type TCPESStats
    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date
End Type

Public TCPESStats As TCPESStats

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function

Public Function VersionesActuales(ByVal v1 As Integer, ByVal v2 As Integer, ByVal v3 As Integer, ByVal v4 As Integer, ByVal v5 As Integer, ByVal v6 As Integer, ByVal v7 As Integer) As Boolean
Dim rv As Boolean

rv = Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "GRAFICOS")) = v1
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "WAVS")) = v2
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MIDIS")) = v3
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "INIT")) = v4
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MAPAS")) = v5
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "AOEXE")) = v6
rv = rv And Val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "EXTRAS")) = v7
VersionesActuales = rv

End Function

Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub


Public Function UnBan(ByVal name As String) As Boolean

End Function

Public Sub BanIpAgrega(ByVal ip As String)
    BanIps.Add ip
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim N As Long

N = BanIpBuscar(ip)
If N > 0 Then
    BanIps.Remove N
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()

End Sub

Public Sub BanIpCargar()

End Sub



Public Sub ActualizaStatsES()

Static TUlt As Single
Dim Transcurrido As Single

Transcurrido = Timer - TUlt

If Transcurrido >= 5 Then
    TUlt = Timer
    With TCPESStats
        .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
        .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
        .BytesEnviados = 0
        .BytesRecibidos = 0
        
        If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
            .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
            .BytesEnviadosXSEGCuando = CDate(Now)
        End If
        
        If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
            .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
            .BytesRecibidosXSEGCuando = CDate(Now)
        End If
        
        If frmEstadisticas.Visible Then
            Call frmEstadisticas.ActualizaStats
        End If
    End With
End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType
'
'Author: Unknown
'03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'
    If EsAdmin(name) Then
        UserDarPrivilegioLevel = PlayerType.admin
    ElseIf EsDios(name) Then
        UserDarPrivilegioLevel = PlayerType.dios
    ElseIf EsSemiDios(name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal reason As String)
    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    Dim bannedIP As String
    With UserList(bannerUserIndex)
        If .admin = True Or .dios = True Then
            tUser = NameIndex(UserName)
            If tUser > 0 Then
                bannedIP = UserList(tUser).ip
            End If
            If LenB(bannedIP) > 0 Then
                Call CloseSocket(tUser)
                Call BanIpAgrega(bannedIP)
            End If
        End If
    End With
End Sub

