Attribute VB_Name = "eRONDAS"
Option Explicit


Public Sub restartround()
        Call SendData(SendTarget.ToAll, 0, PrepareMessageGuildChat("RESTART ROUND EN 1 SEGUNDO..."))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(65, NO_3D_SOUND, NO_3D_SOUND))
        Dim i As Integer

        For i = 1 To maxusers
            With UserList(i)
                If .ConnID <> -1 Then
                    If .ConnIDValida And .flags.UserLogged Then
                                    Call UserDieInterno(i)
                                    Call ResetFrags(i)
                    End If
                End If
            End With
        Next i
        winpk = 0
        winciu = 0
        For i = 1 To LastNPC
            If Npclist(i).bando = eKip.epk Then
                    Npclist(i).Char.body = iCuerpoMuerto
                    Npclist(i).Char.Head = iCabezaMuerto
                    Npclist(i).Char.ShieldAnim = NingunEscudo
                    Npclist(i).Char.WeaponAnim = NingunArma
                    Npclist(i).Char.CascoAnim = NingunCasco
            Else
                    Npclist(i).Char.body = 145
                    Npclist(i).Char.Head = 501
                    Npclist(i).Char.ShieldAnim = NingunEscudo
                    Npclist(i).Char.WeaponAnim = NingunArma
                    Npclist(i).Char.CascoAnim = NingunCasco
            End If
             Call ChangeNPCChar(i, Npclist(i).Char.body, Npclist(i).Char.Head, Npclist(i).Char.heading)
        Next i
        rondax = 0

End Sub

Public Sub dLlevarRand()
        Dim i As Integer
        For i = 1 To maxusers
            With UserList(i)
                If .ConnID <> -1 Then
                    If .ConnIDValida And .flags.UserLogged And .flags.Muerto = 1 And .bando <> eNone Then
                        Llevararand i
                        Call RevivirUsuario1(i)
                    End If
                End If
            End With
        Next i
        rondax = 0
End Sub

Sub Llevararand(UserIndex As Integer)
Dim xx As Byte
Dim yy As Byte
Dim salirfor As Boolean
salirfor = True
Do While salirfor
xx = RandomNumber(10, 85)
yy = RandomNumber(10, 85)
If LegalPos(servermap, xx, yy, False, True) = True And MapData(servermap, xx, yy).UserIndex = 0 And MapData(servermap, xx, yy).NpcIndex = 0 And MapData(servermap, xx, yy).Blocked = 0 Then
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
                        salirfor = False
                        Exit Do
            End If
Loop
RefreshCharStatus UserIndex
UpdateUserInv True, UserIndex, 0
End Sub

Public Sub volverbases()
        Dim i As Integer
        For i = 1 To maxusers
            With UserList(i)
                If .ConnID <> -1 Then
                    If .ConnIDValida And .flags.UserLogged Then
                    volverbase i
                    End If
                End If
            End With
        Next i
        rondax = 0
End Sub

Public Sub volverbase(i As Integer)
            With UserList(i)
                If .ConnID <> -1 Then
                    If .ConnIDValida And .flags.UserLogged Then
                                    LlevaraBase i

                                    If .bando <> eNone Then
                                        Call RevivirUsuario1(i)
                                    End If

                                    UserList(i).ultimomatado = 0
                    End If
                End If
            End With
End Sub





Public Sub cambiarmapa()
        'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(45, 0, 0))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("MAPA CAMBIADO A: " & frmMain.mapax.list(servermap - 1), FontTypeNames.FONTTYPE_TALK))
        Dim i As Integer
        For i = 1 To 100
        If Npclist(i).flags.NPCActive = True Then Call QuitarNPC(i)
        Next i
        For i = 1 To maxusers 'LastUser
            With UserList(i)
               'Conexion activa?
                If .ConnID <> -1 Then
                    '¿User valido?
                    If .ConnIDValida And .flags.UserLogged Then
                        If .bando <> eKip.eNone Then
                            Call UserDieInterno(i)
                            Call RevivirUsuario1(i)
                        End If
                        UserList(i).ultimomatado = 0
                        LlevaraBase i
                    End If
                End If
            End With
        Next i
        'maxusers = MapInfo(servermap).maxusersx
        Call CrearClanPretoriano(50)
        rondax = 0
End Sub
