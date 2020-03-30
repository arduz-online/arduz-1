Attribute VB_Name = "InvUsuario"
Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la funci�n se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(UserIndex).dios = False Then
    If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dej� de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UCase$(MapInfo(UserList(UserIndex).Pos.map).Restringir) = "NEWBIE" Then
    
    Dim DeDonde As WorldPos
    
    Select Case UserList(UserIndex).Hogar
        Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case eCiudad.cUllathorpe
            DeDonde = Ullathorpe
        Case eCiudad.cBanderbill
            DeDonde = Banderbill
        Case Else
            DeDonde = Nix
    End Select
    
    Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .amount = .amount - Cantidad
        '�Quedan mas?
        If .amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount - num

If MapData(map, X, Y).ObjInfo.amount <= 0 Then
    MapData(map, X, Y).ObjInfo.ObjIndex = 0
    MapData(map, X, Y).ObjInfo.amount = 0
    
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount + Obj.amount
    Else
        MapData(map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")

Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'�el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop

'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If

'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
Else
   UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'�Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '�Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.amount = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

            'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
            ElseIf MiObj.amount = 1000 Then 'Es mucha cantidad?
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                    Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name)
                End If
            End If
            
        End If
        
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otAnillo
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    FaccionPuedeUsarItem = True
End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)
        
Select Case Obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.WeaponAnim = NingunArma
                Else
                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
            Else
                UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otAnillo
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                Call DarCuerpoDesnudo(UserIndex, UserList(UserIndex).flags.Mimetizado = 1)
                If Not UserList(UserIndex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.body = Obj.Ropaje
            Else
                UserList(UserIndex).Char.body = Obj.Ropaje
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            UserList(UserIndex).flags.Desnudo = 0
            

        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otCASCO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(UserIndex, Slot)
                If UserList(UserIndex).flags.Mimetizado = 1 Then
                    UserList(UserIndex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.CascoEqpSlot = Slot
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                UserList(UserIndex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
        If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(UserIndex, Slot)
                 If UserList(UserIndex).flags.Mimetizado = 1 Then
                     UserList(UserIndex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(UserIndex).Invent.Object(Slot).Equipped = 1
             UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
             UserList(UserIndex).Invent.EscudoEqpSlot = Slot
             
             If UserList(UserIndex).flags.Mimetizado = 1 Then
                 UserList(UserIndex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                 
                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
             End If
         Else
             Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description)
End Sub

Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

'Verifica si la raza puede usar la ropa
If UserList(UserIndex).raza = eRaza.Humano Or _
   UserList(UserIndex).raza = eRaza.Elfo Or _
   UserList(UserIndex).raza = eRaza.Drow Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If

'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
If (UserList(UserIndex).raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
    CheckRazaUsaRopa = False
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legi�n.
'24/01/2007 Pablo (ToxicWaste) - Utilizaci�n nueva de Barco en lvl 20 por clase Pirata y Pescador.
'

Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(UserIndex).Invent.Object(Slot).amount = 0 Then Exit Sub

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)


If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call WriteConsoleMsg(UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        End If
        
        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(UserIndex).Invent.Object(Slot).Equipped = 0 Then
            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Case eOBJType.otPociones
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteAtacar(UserIndex, False) Then
            Call WriteConsoleMsg(UserIndex, "��Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(UserIndex).flags.TipoPocion
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then _
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            Case 4
                'Usa el item
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, 5)
                If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
       End Select
       Call WriteUpdateUserStats(UserIndex)
       'Call UpdateUserInv(False, UserIndex, Slot)
    Case eOBJType.otInstrumentos
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    Case eOBJType.otBarcos
        If ((LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True, False) _
                Or LegalPos(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True, False)) _
                And UserList(UserIndex).flags.Navegando = 0) _
                Or UserList(UserIndex).flags.Navegando = 1 Then
            Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call WriteConsoleMsg(UserIndex, "�Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
End Select

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteRangingMap(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(UserIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.amount = UserList(UserIndex).Invent.Object(i).amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Gale�n entonces no explota los items. (en el agua)
                If UserList(UserIndex).clase = eClass.Pirat And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
