Attribute VB_Name = "SistemaCombate"
Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18


Function ModificadorEvasion(ByVal clase As eClass) As Single

    ModificadorEvasion = ModClase(clase).Evasion

End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single

    ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
    
    ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

End Function

Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
    
    ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

End Function

Function ModicadorDañoClaseWrestling(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseWrestling = ModClase(clase).DañoWrestling

End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single

    ModEvasionDeEscudoClase = ModClase(clase).Escudo

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function


Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * _
ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
    Dim lTemp As Long
     With UserList(UserIndex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(32 - 12, 0)))
    End With
End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
   PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(40) - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(40) - 12, 0)))

End Function

Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
    PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) * _
    ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
        (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
Else
       PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + _
       (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
End If

PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(40) - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueArma(UserIndex)
    End If
Else 'Peleando con puños
    PoderAtaque = PoderAtaqueWrestling(UserIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
'
Dim rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(UserIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Call WriteBlockedWithShieldUser(UserIndex)

            End If
        End If
    End If
End If
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim DañoMaxArma As Long

''sacar esto si no queremos q la matadracos mate el Dragon si o si
Dim matoDragon As Boolean
matoDragon = False


If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata Dragones?
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
            Else ' Sino es Dragon daño es 1
                DañoArma = 1
                DañoMaxArma = 1
            End If
        Else ' daño comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            DañoArma = 1 ' Si usa la espada mataDragones daño es 1
            DañoMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
                
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                    DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    DañoMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                DañoMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    'Pablo (ToxicWaste)
    ModifClase = ModicadorDañoClaseWrestling(UserList(UserIndex).clase)
    DañoArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
    DañoMaxArma = 3
End If

DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHIT)

''sacar esto si no queremos q la matadracos mate el Dragon si o si
If matoDragon Then
    CalcularDaño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDaño = ((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + DañoUsuario) * ModifClase
End If

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
Dim daño As Long



daño = CalcularDaño(UserIndex, NpcIndex)

'esta navegando? si es asi le sumamos el daño del barco
If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then _
        daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)

daño = daño - Npclist(NpcIndex).Stats.def

If daño < 0 Then daño = 0

'[KEVIN]
Call WriteUserHitNPC(UserIndex, daño)

Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
'[/KEVIN]
UserList(UserIndex).Stats.puntos = UserList(UserIndex).Stats.puntos + (daño / 15)
UserList(UserIndex).Stats.puntosenv = UserList(UserIndex).Stats.puntosenv + (daño / 15)

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apuñalar por la espalda al enemigo
    If PuedeApuñalar(UserIndex) Then
       Call DoApuñalar(UserIndex, NpcIndex, 0, daño)
    End If
    'trata de dar golpe crítico
    Call DoGolpeCritico(UserIndex, NpcIndex, 0, daño)
    
End If

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        
        ' Si era un Dragon perdemos la espada mataDragone
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(UserIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                    Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                    Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                End If
            End If
        Next j
        
        Call MuereNpc(NpcIndex, UserIndex)
End If

End Sub


Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

Dim daño As Integer, Lugar As Integer, absorbido As Integer
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData



daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = daño

If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
    Obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If


Lugar = RandomNumber(1, 6)


Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
           Dim Obj2 As ObjData
           Obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
           If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
           Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           End If
           absorbido = absorbido + defbarco
           daño = daño - absorbido
           If daño < 1 Then daño = 1
        End If
End Select

Call WriteNPCHitUser(UserIndex, Lugar, daño)

UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - daño

If UserList(UserIndex).flags.Meditando Then
    If daño > Fix(UserList(UserIndex).Stats.MinHP / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteMeditateToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    End If
End If

'Muere el usuario
If UserList(UserIndex).Stats.MinHP <= 0 Then

    Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
        End If
    End If
    
    
    Call UserDie(UserIndex)

End If

End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)
    
    Dim EraCriminal As Boolean
    EraCriminal = criminal(UserIndex)
    
    If UserList(UserIndex).Reputacion.BandidoRep > 0 Then
         UserList(UserIndex).Reputacion.BandidoRep = UserList(UserIndex).Reputacion.BandidoRep - vlASALTO
         If UserList(UserIndex).Reputacion.BandidoRep < 0 Then UserList(UserIndex).Reputacion.BandidoRep = 0
    ElseIf UserList(UserIndex).Reputacion.LadronesRep > 0 Then
         UserList(UserIndex).Reputacion.LadronesRep = UserList(UserIndex).Reputacion.LadronesRep - (vlCAZADOR * 10)
         If UserList(UserIndex).Reputacion.LadronesRep < 0 Then UserList(UserIndex).Reputacion.LadronesRep = 0
    End If
    
    If EraCriminal And Not criminal(UserIndex) Then
        Call RefreshCharStatus(UserIndex)
    End If
End Sub


Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
If UserList(UserIndex).bando = Npclist(NpcIndex).bando Then Exit Sub
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
       If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales Or (Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
            If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function

' El npc puede atacar ???
If puede_npc(NpcIndex, 600, False) = True Then
Npclist(NpcIndex).ultimox = GetTickCount()
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, UserIndex, False)

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex

    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And _
       UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
End If
Dim ProbRechazo As Long
If NpcImpacto(NpcIndex, UserIndex) Then
Dim rechazo As Boolean
rechazo = False
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???

          ProbRechazo = Maximo(10, Minimo(90, 50))
          rechazo = (RandomNumber(1, 100) <= ProbRechazo)
    End If
          If rechazo = True Then
          'Se rechazo el ataque con el escudo
                  Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                  Call WriteBlockedWithShieldUser(UserIndex)
          Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            If UserList(UserIndex).flags.Meditando = False Then
                If UserList(UserIndex).flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDaño(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
            

        End If
    End If



'Controla el nivel del usuario
Call CheckUserLevel(UserIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
    Dim daño As Integer
    Dim ANpc As npc
    ANpc = Npclist(Atacante)
    
    daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
    Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - daño
    
    If Npclist(Victima).Stats.MinHP < 1 Then
        
        If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        If Npclist(Atacante).MaestroUser > 0 Then
            Call FollowAmo(Atacante)
        End If
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
    End If
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
If puede_npc(Atacante, 1000, False) = True Then
Npclist(Atacante).ultimox = GetTickCount()
       Npclist(Atacante).CanAttack = 0
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
End If

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)


If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
    Exit Sub
End If

Call NPCAtacado(NpcIndex, UserIndex)

If UserImpactoNpc(UserIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    End If
    
    Call UserDañoNpc(UserIndex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Call WriteUserSwing(UserIndex)
End If

End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

'If UserList(UserIndex).flags.PuedeAtacar = 1 Then
'Check bow's interval
If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub

'Check Spell-Magic interval
If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then
        Exit Sub
    End If
End If

 
'UserList(UserIndex).flags.PuedeAtacar = 0

Dim AttackPos As WorldPos
AttackPos = UserList(UserIndex).Pos
Call HeadtoPos(UserList(UserIndex).Char.heading, AttackPos)
   
'Exit if not legal
If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Exit Sub
End If
    
Dim index As Integer
index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).UserIndex
    
'Look for user
If index > 0 Then
    Call UsuarioAtacaUsuario(UserIndex, index)
    Call WriteUpdateUserStats(UserIndex)
    Call WriteUpdateUserStats(index)
    Exit Sub
End If
    
'Look for NPC
If MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then
    
    If Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).Attackable Then
            
        If Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
            MapInfo(Npclist(MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex).Pos.map).Pk = False Then
                Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
        End If

        Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex)
            
    Else
        Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)
    End If
        
    Call WriteUpdateUserStats(UserIndex)
        
    Exit Sub
End If
    
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
Call WriteUpdateUserStats(UserIndex)


If UserList(UserIndex).Counters.Trabajando Then _
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    
If UserList(UserIndex).Counters.Ocultando Then _
    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

Dim ProbRechazo As Long
Dim rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim UserPoderEvasionEscudo As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim SkillTacticas As Long
Dim SkillDefensa As Long

SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
If Arma > 0 Then
    proyectil = ObjData(Arma).proyectil = 1
Else
    proyectil = False
End If

'Calculamos el poder de evasion...
UserPoderEvasion = PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
   UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
   UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
Else
    UserPoderEvasionEscudo = 0
End If

'Esta usando un arma ???
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = PoderAtaqueArma(AtacanteIndex)
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
   
Else
    PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    ProbExito = Maximo(10, Minimo(90, 50 + _
                ((PoderAtaque - UserPoderEvasion) * 0.4)))
    
End If
UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
    
    'Fallo ???
    If UsuarioImpacto = False Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
      rechazo = (RandomNumber(1, 100) <= ProbRechazo)
      If rechazo = True Then
      'Se rechazo el ataque con el escudo
              Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
              
              Call WriteBlockedWithShieldOther(AtacanteIndex)
              Call WriteBlockedWithShieldUser(VictimaIndex)

      End If
    End If
End If
    
Call FlushBuffer(VictimaIndex)
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   Exit Sub
End If


Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).Pos.X, UserList(AtacanteIndex).Pos.Y))
    
    If UserList(VictimaIndex).flags.Navegando = 0 Then
        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
    End If
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
    
Else
    Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, UserList(AtacanteIndex).Pos.X, UserList(AtacanteIndex).Pos.Y))
    Call WriteUserSwing(AtacanteIndex)
    Call WriteUserAttackedSwing(VictimaIndex, AtacanteIndex)
End If

'If UserList(AtacanteIndex).clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)

End Sub

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim daño As Long, antdaño As Integer
Dim Lugar As Integer, absorbido As Long
Dim defbarco As Integer

Dim Obj As ObjData

daño = CalcularDaño(AtacanteIndex)
antdaño = daño

Call UserEnvenena(AtacanteIndex, VictimaIndex)

If UserList(AtacanteIndex).flags.Navegando = 1 And UserList(AtacanteIndex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim Resist As Byte
If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
    Resist = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo
End If

Lugar = RandomNumber(1, 6)

Select Case Lugar
    Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
        Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        absorbido = absorbido + defbarco - Resist
        daño = daño - absorbido
        If daño < 0 Then daño = 1
        End If
    Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
            Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
            Dim Obj2 As ObjData
            If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
            Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
            absorbido = absorbido + defbarco - Resist
            daño = daño - absorbido
            If daño < 0 Then daño = 1
        End If
End Select

Call WriteUserHittedUser(AtacanteIndex, Lugar, UserList(VictimaIndex).Char.CharIndex, daño)
Call WriteUserHittedByUser(VictimaIndex, Lugar, UserList(AtacanteIndex).Char.CharIndex, daño)

UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - daño

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
        'Trata de apuñalar por la espalda al enemigo
        If PuedeApuñalar(AtacanteIndex) Then
            Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
        End If
        'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
        Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)
End If


If UserList(VictimaIndex).Stats.MinHP <= 0 Then
    'Store it!
    'Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
    
    Call ContarMuerte(VictimaIndex, AtacanteIndex)
    
    ' Para que las mascotas no sigan intentando luchar y
    ' comiencen a seguir al amo
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(AtacanteIndex).MascotasIndex(j) > 0 Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then
                Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
                Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
            End If
        End If
    Next j
    
    Call ActStats(VictimaIndex, AtacanteIndex)
Else
    'Está vivo - Actualizamos el HP
    Call WriteUpdateHP(VictimaIndex)
End If

'Controla el nivel del usuario
Call CheckUserLevel(AtacanteIndex)

Call FlushBuffer(VictimaIndex)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)

    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If criminal(attackerIndex) = criminal(VictimIndex) And atacaequipo = False Then
        Exit Sub
    End If
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(VictimIndex).Char.FX = 0
        UserList(VictimIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
Dim T As eTrigger6
Dim rank As Integer
'MUY importante el orden de estos "IF"...

'Estas muerto no podes atacar
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

'No podes atacar a alguien muerto
If UserList(VictimIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacar = False
    Exit Function
End If

If atacaequipo = False Then
    If criminal(VictimIndex) = criminal(attackerIndex) Then
        Call WriteConsoleMsg(attackerIndex, "No podes atacar a tus compañeros.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
End If

PuedeAtacar = True

End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
'



'Estas muerto?
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Es valida la distancia a la cual estamos atacando?
If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   PuedeAtacarNPC = False
   Exit Function
End If

PuedeAtacarNPC = True

End Function



Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(VictimaIndex).flags.Envenenado = 1
                Call WriteConsoleMsg(VictimaIndex, UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(AtacanteIndex, "Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End If
End If

Call FlushBuffer(VictimaIndex)
End Sub

