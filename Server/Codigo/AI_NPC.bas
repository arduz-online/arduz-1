Attribute VB_Name = "AI"

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

Public Enum e_Alineacion
    ninguna = 0
    Real = 1
    Caos = 2
    Neutro = 3
End Enum

Public Enum e_Personalidad
''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
''Mascota: Solo ataca a quien ataque a su amo.
''Pacifico: No ataca.
    ninguna = 0
    inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5
End Enum

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)
    Dim npos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            npos = .Pos
            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, npos)
                If InMapBounds(npos.map, npos.X, npos.Y) Then
                    UI = MapData(npos.map, npos.X, npos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                            '¿ES CRIMINAL?
                            If Not DelCaos Then
                                If criminal(UI) Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            Else
                                If Not criminal(UI) Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    End If
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
    Dim npos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    
    atacoPJ = False
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            npos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, npos)
                If InMapBounds(npos.map, npos.X, npos.Y) Then
                    UI = MapData(npos.map, npos.X, npos.Y).UserIndex
                    NPCI = MapData(npos.map, npos.X, npos.Y).NpcIndex
                    If UI > 0 And Not atacoPJ Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                            atacoPJ = True
                            If .flags.LanzaSpells <> 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If NpcAtacaUser(NpcIndex, MapData(npos.map, npos.X, npos.Y).UserIndex) Then
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            End If
                            Exit Sub
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
    Dim npos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            npos = .Pos
            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, npos)
                If InMapBounds(npos.map, npos.X, npos.Y) Then
                    UI = MapData(npos.map, npos.X, npos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).name = .flags.AttackedBy Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If UserList(UI).flags.Muerto = 0 Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                        End If
                        
                    End If
                End If
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                        
                    End If
                End If
            Next i
            
            'Si llega aca es que no había ningún usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer

    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select

            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        'If UserList(UI).name = .flags.AttackedBy Then
                        If UserList(UI).bando <> .bando Then
                            If .MaestroUser > 0 Then
                                If criminal(.MaestroUser) = criminal(UI) Then
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NpcIndex, UI)
                                 Else
                                      Call NpcAtacaUser(NpcIndex, UI)
                                 End If
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).bando <> .bando Then
                            If .MaestroUser > 0 Then
                                If criminal(.MaestroUser) = criminal(UI) Then
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NpcIndex, UI)
                                 Else
                                      Call NpcAtacaUser(NpcIndex, UI)
                                 End If
                                 
                                 tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        End If
    End With
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    
    With Npclist(NpcIndex)
        For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
            UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    If Not criminal(UI) Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                    
               End If
            End If
            
        Next i
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If criminal(UI) Then
                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                Exit Sub
                           End If
                        End If
                        
                   End If
                End If
                    
            Next i
        Else
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If criminal(UI) Then
                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                If .flags.Inmovilizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                           End If
                        End If
                        
                   End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim i As Long
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        If .Target = 0 And .TargetNPC = 0 Then
            For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                        If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And UI = .MaestroUser _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                        
                    End If
                End If
                
            Next i
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim X As Long
    Dim Y As Long
    Dim ni As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        ni = MapData(.Pos.map, X, Y).NpcIndex
                        If ni > 0 Then
                            If .TargetNPC = ni Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, ni)
                                    If Npclist(ni).NPCtype = DRAGON Then
                                        Npclist(ni).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(ni, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(ni).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, ni)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                       ni = MapData(.Pos.map, X, Y).NpcIndex
                       If ni > 0 Then
                            If .TargetNPC = ni Then
                                 bNoEsta = True
                                 If .Numero = ELEMENTALFUEGO Then
                                     Call NpcLanzaUnSpellSobreNpc(NpcIndex, ni)
                                     If Npclist(ni).NPCtype = DRAGON Then
                                        Npclist(ni).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(ni, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(ni).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, ni)
                                    End If
                                 End If
                                 If .flags.Inmovilizado = 1 Then Exit Sub
                                 If .TargetNPC = 0 Then Exit Sub
                                 tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, X, Y).NpcIndex).Pos)
                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
    With Npclist(NpcIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If .NPCtype = eNPCType.GuardiaReal Then
                Call GuardiasAI(NpcIndex, False)
            ElseIf .NPCtype = eNPCType.Guardiascaos Then
                Call GuardiasAI(NpcIndex, True)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement
            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Then Exit Sub
                If .NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf .NPCtype = eNPCType.Guardiascaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                End If
            
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
    UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    tmpPos.map = Npclist(NpcIndex).Pos.map
    tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y ' invertí las coordenadas
    tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X
    
    'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
    
    tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)
    
    MoveNPCChar NpcIndex, tHeading
    
    Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1
End Function


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
    
    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
End Sub
