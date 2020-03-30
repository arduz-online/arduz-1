Attribute VB_Name = "aCOSAS"
Option Explicit
Public svname As String
Public menduz As String
Public valeresu As Boolean
Public valeestu As Boolean
Public valeinvi As Boolean
Public rondaa As Long
Public rondaact As Boolean
Public serverrunning As Boolean
Public servermap As Integer
Public atacaequipo As Boolean
Public enviarank As Boolean
Public deathm As Boolean
Public botsact As Boolean
Public mankismo As Integer
Public winpk As Long
Public winciu As Long
Public fatuos As Boolean

Public hIP As String

Public WEBCLASS As clsMenduz

Public adminpasswd As String

Public deathmatch As Boolean

Type jugador
    UserIndex As Integer
    Activado As Boolean
    Frags As Integer
    gano As Integer
    muertes As Integer
End Type

Type ekipos
    Jugadores(1 To 50) As jugador
    NumJugadores As Integer
    gano As Integer
    perdio As Integer
    npcact As Boolean
    NPCs(1 To 50) As Integer
End Type

Public equipos(1 To 2) As ekipos

Public Const WEBSERVER As String = "http://ao.noicoder.com/" '"http://localhost/ao/"


Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       


Public Sub enviaser()
Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Servidor: " & svname & " - Tiempo de ronda: " & IIf(rondaact = True, (rondaa / 60) & " Minutos.", "Infinito")))
If valeinvi = True Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Invisibilidad esta ACTIVADA"))
Else
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Invisibilidad esta DESACTIVADA"))
End If
If valeestu = True Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Estupidez esta ACTIVADA"))
Else
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Estupidez esta DESACTIVADA"))
End If
If valeresu = True Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Resucitar esta ACTIVADA"))
Else
    Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("Resucitar esta DESACTIVADA"))
End If
End Sub

Public Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    Call WriteGuildChat(UserIndex, "¡Bienvenido a Arduz!")
End Sub


Public Sub enviaser1(UI As Integer)
Call WriteGuildChat(UI, "Servidor: " & svname & " - Tiempo de ronda: " & IIf(rondaact = True, (rondaa / 60) & " Minutos.", "Infinito"))
If valeinvi = True Then
    Call WriteGuildChat(UI, "Invisibilidad esta ACTIVADA")
Else
    Call WriteGuildChat(UI, "Invisibilidad esta DESACTIVADA")
End If
If valeestu = True Then
    Call WriteGuildChat(UI, "Estupidez esta ACTIVADA")
Else
    Call WriteGuildChat(UI, "Estupidez esta DESACTIVADA")
End If
If valeresu = True Then
    Call WriteGuildChat(UI, "Resucitar esta ACTIVADA")
Else
    Call WriteGuildChat(UI, "Resucitar esta DESACTIVADA")
End If
End Sub


'////////////////////////////////////////////////
Public Function UserVivos(Tipo As eKip) As Integer
Dim total As Integer
Dim i As Integer
        For i = 1 To maxusers 'LastUser
            With UserList(i)
               'Conexion activa?
                If .ConnID <> -1 Then
                    '¿User valido?
                    If .ConnIDValida And .flags.UserLogged And .bando = Tipo And .flags.Muerto = 0 Then
                        total = total + 1
                    End If
                End If
            End With
        Next i
UserVivos = total
End Function

Public Function UserMuertos(Tipo As eKip) As Integer
Dim total As Integer
Dim i As Integer
        For i = 1 To maxusers 'LastUser
            With UserList(i)
               'Conexion activa?
                If .ConnID <> -1 Then
                    '¿User valido?
                    If .ConnIDValida And .flags.UserLogged And .bando = Tipo And .flags.Muerto = 1 Then
                        total = total + 1
                    End If
                End If
            End With
        Next i
UserMuertos = total
End Function

Public Function UserBando(Tipo As eKip) As Integer
Dim total As Integer
Dim i As Integer
        For i = 1 To maxusers 'LastUser
        maxusers = 20
            With UserList(i)
               'Conexion activa?
                If .ConnID <> -1 Then
                    '¿User valido?
                    If .ConnIDValida And .flags.UserLogged And .bando = Tipo Then
                        total = total + 1
                    End If
                End If
            End With
        Next i
UserBando = total
End Function

Public Function BalancearPJ(useridex As Integer) As Integer
'Dim lvl As Integer
'Dim lvlm As Integer
'Dim i As Integer
            'With UserList(useridex)
            '    If .ConnID <> -1 Then
            '        If .ConnIDValida And .flags.UserLogged Then
            '        'If .ConnIDValida And .flags.UserLogged And IIf(Tipo = eKip.eNone, Tipo, .bando) = Tipo And IIf(UID <> 0, UID, i) <> i Then
            '            If .Stats.ELV < MaxLvlGet(, useridex) Then
            '                lvlm = MaxLvlGet(, useridex) - .Stats.ELV
            '                For i = 0 To lvlm
            '                    SubirLVLPj useridex
            '                    Debug.Print "ASD"
            '                Next i
            '            End If
            '        End If
            '    End If
            'End With
            'Call WriteUpdateUserStats(useridex)
End Function

Sub LlevaraTrigger(trigger As eTrigger, UserIndex As Integer, Optional warp As Boolean = True)
Dim xx As Byte
Dim yy As Byte
Dim salirfor As Boolean
For xx = 9 To 90
    If salirfor = False Then
        For yy = 9 To 90
            If MapData(servermap, xx, yy).trigger = trigger And LegalPos(servermap, xx, yy, False, True) = True And (MapData(servermap, xx, yy).UserIndex <> 0 Or MapData(servermap, xx, yy).NpcIndex <> 0) Then
                        If warp = False Then
                            UserList(UserIndex).Pos.X = xx
                            UserList(UserIndex).Pos.Y = yy
                        Else
                            Call WarpUserChar(UserIndex, servermap, xx, yy)
                            If UserList(UserIndex).flags.Paralizado = 1 Then
                                UserList(UserIndex).flags.Paralizado = 0
                                Call WriteParalizeOK(UserIndex)
                            End If
                            
                            '<<< Estupidez >>>
                            If UserList(UserIndex).flags.Estupidez = 1 Then
                                UserList(UserIndex).flags.Estupidez = 0
                                Call WriteDumbNoMore(UserIndex)
                            End If
                            salirfor = True
                        End If
                        Exit For
            End If
        Next yy
    Else
        Exit For
    End If
Next xx

End Sub

Sub LlevaraBase(UserIndex As Integer)
Dim xx As Byte
Dim yy As Byte
Dim trigger As eTrigger
If UserList(UserIndex).bando = eKip.epk Then
trigger = eTrigger.RESUPK
ElseIf UserList(UserIndex).bando = eKip.eCUI Then
trigger = eTrigger.RESUCIU
Else
Llevararand UserIndex
Exit Sub
End If
Dim salirfor As Boolean
For xx = 9 To 90
    If salirfor = False Then
        For yy = 9 To 90
            If MapData(servermap, xx, yy).trigger = trigger And LegalPos(servermap, xx, yy, False, True) = True And MapData(servermap, xx, yy).UserIndex = 0 And MapData(servermap, xx, yy).NpcIndex = 0 Then
                        Call WarpUserChar(UserIndex, servermap, xx, yy)
                        If UserList(UserIndex).flags.Paralizado = 1 Then
                            UserList(UserIndex).flags.Paralizado = 0
                            Call WriteParalizeOK(UserIndex)
                        End If
                        
                        '<<< Estupidez >>>
                        If UserList(UserIndex).flags.Estupidez = 1 Then
                            UserList(UserIndex).flags.Estupidez = 0
                            Call WriteDumbNoMore(UserIndex)
                        End If
                        salirfor = True
                        Exit For
            End If
        Next yy
    Else
        Exit For
    End If
Next xx

RefreshCharStatus UserIndex
'UpdateUserInv True, UserIndex, 0
End Sub

Public Function puede_npc(i As Integer, intervalo As Long, Optional modif As Boolean = True) As Boolean
Dim tmp As Boolean
tmp = (GetTickCount() - (Npclist(i).ultimox + intervalo)) > -1
'Debug.Print (GetTickCount() - (Npclist(i).ultimox + intervalo))
If tmp = True Then
Npclist(i).ultimox = GetTickCount()
End If
puede_npc = tmp

End Function

' Convert a zero-terminated fixed string to a dynamic VB string
Public Function sz2string(ByVal szStr As String) As String
    sz2string = Left$(szStr, InStr(1, szStr, Chr$(0)) - 1)
End Function
