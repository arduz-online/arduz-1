Attribute VB_Name = "Trabajo"
Option Explicit

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'

UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1
If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
    UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(UserIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(UserIndex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(UserIndex).Counters.TiempoOculto = 0
    UserList(UserIndex).flags.Oculto = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
End If



Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo Errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(UserIndex).flags.Oculto = 1
    Suerte = (-0.000001 * (100 - Skill) ^ 3)
    Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
    Suerte = Suerte + (-0.0088 * (100 - Skill))
    Suerte = Suerte + (0.9571)
    Suerte = Suerte * IntervaloOculto
    UserList(UserIndex).Counters.TiempoOculto = Suerte
  
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(UserIndex).clase)

If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.BarcoSlot = Slot

If UserList(UserIndex).flags.Navegando = 0 Then
    
    UserList(UserIndex).Char.Head = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        '(Nacho)
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
    Else
        UserList(UserIndex).Char.body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
    
Else
    
    UserList(UserIndex).flags.Navegando = 0
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
        
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteNavigateToggle(UserIndex)

End Sub


Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        total = total + UserList(UserIndex).Invent.Object(i).amount
    End If
Next i

If cant <= total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(UserIndex, i)
        
        UserList(UserIndex).Invent.Object(i).amount = UserList(UserIndex).Invent.Object(i).amount - cant
        If (UserList(UserIndex).Invent.Object(i).amount <= 0) Then
            cant = Abs(UserList(UserIndex).Invent.Object(i).amount)
            UserList(UserIndex).Invent.Object(i).amount = 0
            UserList(UserIndex).Invent.Object(i).ObjIndex = 0
        Else
            cant = 0
        End If
        
        Call UpdateUserInv(False, UserIndex, i)
        
        If (cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Function ModNavegacion(ByVal clase As eClass) As Single
ModNavegacion = 2.3
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
            UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
            UserList(UserIndex).Counters.Mimetismo = 0
            UserList(UserIndex).flags.Mimetizado = 0
        End If
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Char.body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
End Sub



Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'
'Autor: Nacho (Integer) & Unknown (orginal version)
'04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'
Dim Suerte As Integer
Dim Skill As Integer

Skill = 100

Select Case UserList(UserIndex).clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    Case eClass.Cleric, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).clase = eClass.Assasin Then
            daño = Round(daño * 1.4, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
    End If
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'
'Autor: Pablo (ToxicWaste)
'28/01/2007
'
Dim Suerte As Integer
Dim Skill As Integer

If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.5)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]

    End If
End If

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If
    
If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMeditateToggle(UserIndex)
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    Exit Sub
End If
Suerte = 5

res = RandomNumber(1, Suerte)

If res = 1 Then
    cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, 3)
    If cant <= 0 Then cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call WriteUpdateMana(UserIndex)
End If

End Sub
