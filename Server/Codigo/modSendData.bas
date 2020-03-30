Attribute VB_Name = "modSendData"

Option Explicit

Public Enum SendTarget
    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
'

'Last Modify Date: 01/08/2007
'Last modified by: (liquid)
'
On Error Resume Next
    Dim LoopC As Long
    Dim map As Integer
    
    Select Case sndRoute
        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).admin = True Or UserList(LoopC).dios = True Then
                        Call EnviarDatosASlot(LoopC, sndData)
                   End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAll
            For LoopC = 1 To LastUser
                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToAllButIndex
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToConsejo
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToConsejoCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRolesMasters
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminales
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToReal
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaos
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToRealYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
        
        Case SendTarget.ToCaosYRMs
            For LoopC = 1 To LastUser
                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)
                    End If
                End If
            Next LoopC
            Exit Sub
    End Select
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)
'
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If tempIndex <> UserIndex Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
'

'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True Then 'And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.admin Or PlayerType.dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
'

'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    If UserList(UserIndex).guildIndex = 0 Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).guildIndex = UserList(UserIndex).guildIndex Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)
'

'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
        
        If UserList(UserIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(UserIndex).AreasInfo.AreaReciveY And AreaY Then
                    Call EnviarDatosASlot(UserIndex, sdData)
            End If
        End If
    'Next LoopC
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String)
'

'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = UserList(UserIndex).Pos.map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.dios Or PlayerType.admin) Then _
                        Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
'
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    Dim map As Integer
    Dim AreaX As Integer
    Dim AreaY As Integer
    
    map = Npclist(NpcIndex).Pos.map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Public Sub SendToAreaByPos(ByVal map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
'
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim tempIndex As Integer
    
    AreaX = 2 ^ (AreaX \ 9)
    AreaY = 2 ^ (AreaY \ 9)
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)
                End If
            End If
        End If
    Next LoopC
End Sub

Public Sub SendToMap(ByVal map As Integer, ByVal sdData As String)
'

'Last Modify Date: 5/24/2007
'
'
    Dim LoopC As Long
    Dim tempIndex As Integer
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next LoopC
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
'

'Last Modify Date: 5/24/2007
'
'
    Dim LoopC As Long
    Dim map As Integer
    Dim tempIndex As Integer
    
    map = UserList(UserIndex).Pos.map
    
    If Not MapaValido(map) Then Exit Sub

    For LoopC = 1 To ConnGroups(map).CountEntrys
        tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)
        End If
    Next LoopC
End Sub
