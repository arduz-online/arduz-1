Attribute VB_Name = "UsUaRiOs"
Option Explicit

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)

Dim DaExp As Long
Dim EraCriminal As Boolean

DaExp = UserList(VictimIndex).Stats.ELV * 500

UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp
If UserList(attackerIndex).Stats.Exp > MAXEXP Then _
    UserList(attackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)

Call UserDie(VictimIndex)

If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then _
    UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1

Call FlushBuffer(VictimIndex)

'Log

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
If UserList(UserIndex).bando = eKip.eNone Then Exit Sub
UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 10

'No puede estar empollando
UserList(UserIndex).flags.EstaEmpo = 0
UserList(UserIndex).EmpoCont = 0

If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
End If

If UserList(UserIndex).flags.Navegando = 1 Then
    Dim Barco As ObjData
    Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
        UserList(UserIndex).Char.body = iFragataReal
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
        UserList(UserIndex).Char.body = iFragataCaos
    Else
        If criminal(UserIndex) Then
            If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaPk
            If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraPk
            If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonPk
        Else
            If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
            If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
            If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
        End If
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco

Else
    Call DarCuerpoDesnudo(UserIndex)
    
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
End If



Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteUpdateUserStats(UserIndex)

End Sub


Sub RevivirUsuario1(ByVal UserIndex As Integer)
'If UserList(UserIndex).bando = eKip.eNone Then Exit Sub
If UserList(UserIndex).bando <> eNone Then
Dim ipa As Integer
ipa = UserList(UserIndex).flags.Muerto
    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    If UserList(UserIndex).flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)
        UserList(UserIndex).flags.Paralizado = 0
        UserList(UserIndex).flags.Inmovilizado = 0
    End If

    'No puede estar empollando
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
    
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then
        UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    End If

    If UserList(UserIndex).flags.Navegando Then
        Dim Barco As ObjData
        Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
        UserList(UserIndex).Char.Head = 0
            
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            UserList(UserIndex).Char.body = iFragataReal
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            UserList(UserIndex).Char.body = iFragataCaos
        Else
            If criminal(UserIndex) Then
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaPk
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraPk
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonPk
            Else
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
            End If
        End If
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        If ipa = 1 Then
           Call DarCuerpoDesnudo(UserIndex)
        End If
    End If
    UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
    

Else
Llevararand UserIndex
End If
Call WriteMiniStats(UserIndex)
Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteUpdateUserStats(UserIndex)

End Sub


Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer)

    With UserList(UserIndex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
    End With
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, Head, heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, casco))
End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
    Dim L As Long
    
    L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
        (-UserList(UserIndex).Reputacion.BandidoRep) + _
        UserList(UserIndex).Reputacion.BurguesRep + _
        (-UserList(UserIndex).Reputacion.LadronesRep) + _
        UserList(UserIndex).Reputacion.NobleRep + _
        UserList(UserIndex).Reputacion.PlebeRep
    L = Round(L / 6)
    
    UserList(UserIndex).Reputacion.Promedio = L
    
    Call WriteFame(UserIndex)
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex))
    Call QuitarUser(UserIndex, UserList(UserIndex).Pos.map)
    
    MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
    Dim klan As String
    Dim Barco As ObjData
    
    If Len(UserList(UserIndex).modName) > 0 Then
        klan = UserList(UserIndex).modName
        klan = " <" & klan & ">"
    End If
    
    If UserList(UserIndex).showName And UserList(UserIndex).bando <> eNone Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, criminal(UserIndex), UserList(UserIndex).name & klan))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, criminal(UserIndex), vbNullString))
    End If
    
    'Si esta navengando, se cambia la barca.
    If UserList(UserIndex).flags.Navegando Then
        Barco = ObjData(UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex)
        
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            UserList(UserIndex).Char.body = iFragataReal
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
            UserList(UserIndex).Char.body = iFragataCaos
        Else
            If criminal(UserIndex) Then
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaPk
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraPk
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonPk
            Else
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
            End If
        End If
        
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    End If
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(map, X, Y) Then
        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If
        
        'Place character on map if needed
        If toMap Then _
            MapData(map, X, Y).UserIndex = UserIndex
        
        'Send make character command to clients
        Dim klan As String
        If Len(UserList(UserIndex).modName) > 0 Then
            klan = UserList(UserIndex).modName
        End If
        
        Dim bCr As Byte
        
        bCr = criminal(UserIndex)
        
        If LenB(klan) <> 0 Then
            If Not toMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name & " <" & klan & ">", bCr, color_nick_user(UserIndex))
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        Else 'if tiene clan
            If Not toMap Then
                If UserList(UserIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name, bCr, color_nick_user(UserIndex))
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User)
                End If
            Else
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.map)
            End If
        End If 'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(UserIndex)
End Sub




Function color_nick_user(UI As Integer)
If (UserList(UI).bando = eKip.eNone) Then
    color_nick_user = 8
Else
    If deathm = False Then
        If UserList(UI).dios = True Then
            color_nick_user = 20
        Else
            color_nick_user = UserList(UI).flags.Privilegios
        End If
    Else
       color_nick_user = 15
    End If
End If
End Function




Sub CheckUserLevel(ByVal UserIndex As Integer)

End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

Dim npos As WorldPos
Dim sailing As Boolean


    sailing = PuedeAtravesarAgua(UserIndex)
    npos = UserList(UserIndex).Pos
    Call HeadtoPos(nHeading, npos)
    
    If LegalPos(UserList(UserIndex).Pos.map, npos.X, npos.Y, sailing, Not sailing) Then
        If MapInfo(UserList(UserIndex).Pos.map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, npos.X, npos.Y))

        End If
        
        'Update map and user pos
        MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = npos
        UserList(UserIndex).Char.heading = nHeading
        MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
    Else
        Call WritePosUpdate(UserIndex)
    End If
    
    If UserList(UserIndex).Counters.Trabajando Then _
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1

    If UserList(UserIndex).Counters.Ocultando Then _
        UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To maxusers + 1
        If LoopC > maxusers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > maxusers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > maxusers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
End If

End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
Dim EraCriminal As Boolean

'Guardamos el usuario que ataco el npc.
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).name

'Npc que estabas atacando.
Dim LastNpcHit As Integer
LastNpcHit = UserList(UserIndex).flags.NPCAtacado
'Guarda el NPC que estas atacando ahora.
UserList(UserIndex).flags.NPCAtacado = NpcIndex

If Npclist(NpcIndex).MaestroUser > 0 Then
    If Npclist(NpcIndex).MaestroUser <> UserIndex Then
        Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
    End If
End If
    
    If Npclist(NpcIndex).MaestroUser <> UserIndex Then
        'hacemos que el npc se defienda
        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1
    End If


End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(UserIndex).clase = eClass.Assasin) And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function


''
' Muere un usuario
'
'UserIndex  Indice del usuario que muere
'



Sub UserDieInterno(ByVal UserIndex As Integer)
'*
'Author: Uknown
'Last Modified: 04/15/2008 (NicoNZ)
'Ahora se resetea el counter del invi
'*
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1
    UserList(UserIndex).flags.SeguroResu = False
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    aN = UserList(UserIndex).flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)
    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call WriteRestOK(UserIndex)
    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
    End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
    End If
    'desequipar municiones
    If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).Char.loops = INFINITE_LOOPS Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
        UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
    End If
    
    ' << Restauramos los atributos >>
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 18 + ModRaza(UserRaza).Fuerza
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 18 + ModRaza(UserRaza).Agilidad
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = 18 + ModRaza(UserList(UserIndex).raza).Inteligencia
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = 18 + ModRaza(UserList(UserIndex).raza).Carisma
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = 18 + ModRaza(UserList(UserIndex).raza).Constitucion
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).bando = eKip.epk Then
        If UserList(UserIndex).flags.Navegando = 0 Then
            UserList(UserIndex).Char.body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        Else
            UserList(UserIndex).Char.body = iFragataFantasmal ';)
        End If
    Else
        If UserList(UserIndex).flags.Navegando = 0 Then
            UserList(UserIndex).Char.body = 145
            UserList(UserIndex).Char.Head = 501
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        Else
            UserList(UserIndex).Char.body = iFragataFantasmal ';)
        End If
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    Next i
    
    UserList(UserIndex).NroMascotas = 0
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(UserIndex)

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub




Sub UserDie(ByVal UserIndex As Integer)
'*
'Author: Uknown
'Last Modified: 04/15/2008 (NicoNZ)
'Ahora se resetea el counter del invi
'*
On Error GoTo ErrorHandler
    Dim i As Long
    Dim aN As Integer
    
    'Sonido
    If UserList(UserIndex).genero = eGenero.Mujer Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    UserList(UserIndex).Stats.MinHP = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Muerto = 1
    UserList(UserIndex).flags.SeguroResu = False
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString
    End If
    
    aN = UserList(UserIndex).flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)
    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)
    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call WriteRestOK(UserIndex)
    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
    End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura

        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
        End If
        'desequipar arma
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
        End If
        'desequipar casco
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
        End If
        'desequipar herramienta
        If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
        End If
        'desequipar municiones
        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
        End If
        'desequipar escudo
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
        End If

    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).Char.loops = INFINITE_LOOPS Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(UserIndex).flags.Mimetizado = 1 Then
        UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
        UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
        UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
        UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
        UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
        UserList(UserIndex).Counters.Mimetismo = 0
        UserList(UserIndex).flags.Mimetizado = 0
    End If

    ' << Restauramos los atributos >>
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 18 + ModRaza(UserRaza).Fuerza
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 18 + ModRaza(UserRaza).Agilidad
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = 18 + ModRaza(UserList(UserIndex).raza).Inteligencia
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = 18 + ModRaza(UserList(UserIndex).raza).Carisma
        'UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = 18 + ModRaza(UserList(UserIndex).raza).Constitucion
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).bando = eKip.epk Then
        If UserList(UserIndex).flags.Navegando = 0 Then
            UserList(UserIndex).Char.body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        Else
            UserList(UserIndex).Char.body = iFragataFantasmal ';)
        End If
    Else
        If UserList(UserIndex).flags.Navegando = 0 Then
            UserList(UserIndex).Char.body = 145
            UserList(UserIndex).Char.Head = 501
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco
        Else
            UserList(UserIndex).Char.body = iFragataFantasmal ';)
        End If
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
        End If
    Next i
    
    UserList(UserIndex).NroMascotas = 0
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(UserIndex)

Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
If UserList(Atacante).ultimomatado <> Muerto Then
    UserList(Atacante).Stats.puntos = UserList(Atacante).Stats.puntos + 150
    UserList(Atacante).Stats.puntosenv = UserList(Atacante).Stats.puntosenv + 150
    UserList(Atacante).ultimomatado = Muerto
End If
UserList(Atacante).Stats.UsuariosMatadosenv = UserList(Atacante).Stats.UsuariosMatadosenv + 1
UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
UserList(Muerto).Stats.muertes = UserList(Muerto).Stats.muertes + 1
UserList(Muerto).Stats.muertesenv = UserList(Muerto).Stats.muertesenv + 1
End Sub

Public Sub ResetFrags(ByVal UI As Integer)
UserList(UI).Faccion.CiudadanosMatados = 0
UserList(UI).Stats.UsuariosMatados = 0
UserList(UI).Stats.UsuariosMatadosenv = 0
UserList(UI).Stats.muertes = 0
UserList(UI).Stats.muertesenv = 0
UserList(UI).Stats.puntos = 0
UserList(UI).Stats.puntosenv = 0
UserList(UI).ultimomatado = 0
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef npos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    npos.map = Pos.map
    
    Do While Not LegalPos(Pos.map, npos.X, npos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(npos.map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(npos.map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(npos.map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(npos.map, tX, tY).ObjInfo.amount + Obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(npos.map, tX, tY).TileExit.map = 0 Then
                        npos.X = tX
                        npos.Y = tY
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        npos.X = 0
        npos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    'Quitar el dialogo
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    Call WriteRemoveAllDialogs(UserIndex)
    
    OldMap = UserList(UserIndex).Pos.map
    OldX = UserList(UserIndex).Pos.X
    OldY = UserList(UserIndex).Pos.Y
    
    Call EraseUserChar(UserIndex)
    
    If OldMap <> map Then
        Call WriteChangeMap(UserIndex, map, MapInfo(UserList(UserIndex).Pos.map).MapVersion)
        Call WritePlayMidi(UserIndex, Val(ReadField(1, MapInfo(map).Music, 45)))
        
        'Update new Map Users
        MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
        
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.map = map
    
    Call MakeUserChar(True, map, UserIndex, map, X, Y)
    Call WriteUserCharIndexInServer(UserIndex)
    
    'Force a flush, so user index is in there before it's destroyed for teleporting
    Call FlushBuffer(UserIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
    End If
    
    If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    Call WarpMascotas(UserIndex)
End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(UserIndex).NroMascotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
                UserList(UserIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
        ElseIf UserList(UserIndex).MascotasType(i) > 0 Then
            PetRespawn(i) = True
            PetTypes(i) = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida(i) = 0
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
          If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
            'Controlamos que se sumoneo OK
            If UserList(UserIndex).MascotasIndex(i) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
                Exit For
            End If
            Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
            Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
          Else
            Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
            Exit For
          End If
        End If
    Next i
    
    UserList(UserIndex).NroMascotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(UserIndex).NroMascotas Then UserList(UserIndex).NroMascotas = 0

End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer)

    Dim isNotVisible As Boolean
    
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = 0
        
        isNotVisible = (UserList(UserIndex).flags.Oculto Or UserList(UserIndex).flags.invisible)
        If isNotVisible Then
            UserList(UserIndex).flags.Oculto = 0
            UserList(UserIndex).flags.invisible = 0
            UserList(UserIndex).Counters.Invisibilidad = 0
            UserList(UserIndex).Counters.TiempoOculto = 0
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
        End If
    End If
End Sub

Public Sub CancelExit(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Saliendo Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = 1
        End If
    End If
End Sub


Public Sub Empollando(ByVal UserIndex As Integer)
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    UserList(UserIndex).flags.EstaEmpo = 1
Else
    UserList(UserIndex).flags.EstaEmpo = 0
    UserList(UserIndex).EmpoCont = 0
End If
End Sub

Public Function BodyIsBoat(ByVal body As Integer) As Boolean

    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function
