Attribute VB_Name = "ES"
Option Explicit



Function EsAdmin(ByVal name As String) As Boolean

End Function

Function EsDios(ByVal name As String) As Boolean

End Function

Function EsSemiDios(ByVal name As String) As Boolean

End Function

Function EsConsejero(ByVal name As String) As Boolean

End Function

Function EsRolesMaster(ByVal name As String) As Boolean

End Function


Public Function TxtDimension(ByVal name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()

ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub

Public Sub CargarHechizos()
On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))

ReDim Hechizos(1 To NumeroHechizos) As tHechizo

'frmCargando.cargar.min = 0
'frmCargando.cargar.max = NumeroHechizos
'frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = Val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = Val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
'    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = Val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = Val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = Val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = Val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = Val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = Val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = Val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).cant = Val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = Val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    
'    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
'    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = Val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    'frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    Hechizos(Hechizo).NeedStaff = Val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim LoopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(map, X, Y).Graphic(1)
                
                For LoopC = 2 To 4
                    If MapData(map, X, Y).Graphic(LoopC) Then _
                        Put FreeFileMap, , MapData(map, X, Y).Graphic(LoopC)
                Next LoopC
                
                If MapData(map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
                   If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(map, X, Y).ObjInfo.ObjIndex = 0
                        MapData(map, X, Y).ObjInfo.amount = 0
                    End If
                End If
    
                If MapData(map, X, Y).TileExit.map Then ByFlags = ByFlags Or 1
                If MapData(map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(map, X, Y).ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(map, X, Y).TileExit.map Then
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.map
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(map, X, Y).TileExit.Y
                End If
                
                If MapData(map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(map, X, Y).NpcIndex).Numero
                
                If MapData(map, X, Y).ObjInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(map, X, Y).ObjInfo.ObjIndex
                    Put FreeFileInf, , MapData(map, X, Y).ObjInfo.amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Name", MapInfo(map).name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "MusicNum", MapInfo(map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "MagiaSinefecto", MapInfo(map).MagiaSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "InviSinEfecto", MapInfo(map).InviSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "mapa" & map, "ResuSinEfecto", MapInfo(map).ResuSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "StartPos", MapInfo(map).StartPos.map & "-" & MapInfo(map).StartPos.X & "-" & MapInfo(map).StartPos.Y)
    

    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Terreno", MapInfo(map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Zona", MapInfo(map).Zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Restringir", MapInfo(map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "BackUp", str(MapInfo(map).BackUp))

    If MapInfo(map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()


End Sub

Sub LoadArmadurasHerreria()


End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    'For i = 1 To NUMCLASES
    '    ModClase(i).Evasion = Val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
    '    ModClase(i).AtaqueArmas = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
    '    ModClase(i).AtaqueProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
    '    ModClase(i).DañoArmas = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
    '    ModClase(i).DañoProyectiles = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
    '    ModClase(i).DañoWrestling = Val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
    '    ModClase(i).Escudo = Val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
    'Next i
        i = eClass.Warrior
        ModClase(i).Evasion = 1
        ModClase(i).AtaqueArmas = 1
        ModClase(i).AtaqueProyectiles = 0.8
        ModClase(i).DañoArmas = 1.1
        ModClase(i).DañoProyectiles = 1.08
        ModClase(i).DañoWrestling = 0.1
        ModClase(i).Escudo = 0.8
        i = eClass.Hunter
        ModClase(i).Evasion = 0.9
        ModClase(i).AtaqueArmas = 0.8
        ModClase(i).AtaqueProyectiles = 1
        ModClase(i).DañoArmas = 0.9
        ModClase(i).DañoProyectiles = 1.1
        ModClase(i).Escudo = 0.72
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Paladin
        ModClase(i).Evasion = 0.85
        ModClase(i).AtaqueArmas = 0.85
        ModClase(i).AtaqueProyectiles = 0.75
        ModClase(i).DañoArmas = 0.9
        ModClase(i).DañoProyectiles = 0.8
        ModClase(i).Escudo = 1
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Assasin
        ModClase(i).Evasion = 1.1
        ModClase(i).AtaqueArmas = 0.85
        ModClase(i).AtaqueProyectiles = 0.75
        ModClase(i).DañoArmas = 0.9
        ModClase(i).DañoProyectiles = 0.8
        ModClase(i).Escudo = 0.7
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Bard
        ModClase(i).Evasion = 1.1
        ModClase(i).AtaqueArmas = 0.75
        ModClase(i).AtaqueProyectiles = 0.7
        ModClase(i).DañoArmas = 0.75
        ModClase(i).DañoProyectiles = 0.7
        ModClase(i).Escudo = 0.65
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Cleric
        ModClase(i).Evasion = 0.81
        ModClase(i).AtaqueArmas = 0.85
        ModClase(i).AtaqueProyectiles = 0.7
        ModClase(i).DañoArmas = 0.85
        ModClase(i).DañoProyectiles = 0.7
        ModClase(i).Escudo = 0.8
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Druid
        ModClase(i).Evasion = 0.85
        ModClase(i).AtaqueArmas = 0.6
        ModClase(i).AtaqueProyectiles = 0.7
        ModClase(i).DañoArmas = 0.7
        ModClase(i).DañoProyectiles = 0.7
        ModClase(i).Escudo = 0.6
        ModClase(i).DañoWrestling = 0.1
        i = eClass.Mage
        ModClase(i).Evasion = 0.7
        ModClase(i).AtaqueArmas = 0.5
        ModClase(i).AtaqueProyectiles = 0.5
        ModClase(i).DañoArmas = 0.5
        ModClase(i).DañoProyectiles = 0.6
        ModClase(i).Escudo = 0.6
        ModClase(i).DañoWrestling = 0.1
    'Modificadores de Raza
    
    ModRaza(eRaza.Humano).Fuerza = 20
    ModRaza(eRaza.Humano).Agilidad = 20
    ModRaza(eRaza.Humano).Inteligencia = 0
    ModRaza(eRaza.Humano).Carisma = 0
    ModRaza(eRaza.Humano).Constitucion = 2
    
    ModRaza(eRaza.Drow).Fuerza = 20
    ModRaza(eRaza.Drow).Agilidad = 20
    ModRaza(eRaza.Drow).Inteligencia = 2
    ModRaza(eRaza.Drow).Carisma = -3
    ModRaza(eRaza.Drow).Constitucion = 1
    
    ModRaza(eRaza.Elfo).Fuerza = 36 - 18
    ModRaza(eRaza.Elfo).Agilidad = 20
    ModRaza(eRaza.Elfo).Inteligencia = 2
    ModRaza(eRaza.Elfo).Carisma = 2
    ModRaza(eRaza.Elfo).Constitucion = 1
    
    ModRaza(eRaza.Enano).Fuerza = 20
    ModRaza(eRaza.Enano).Agilidad = 34 - 18
    ModRaza(eRaza.Enano).Inteligencia = -6
    ModRaza(eRaza.Enano).Carisma = -2
    ModRaza(eRaza.Enano).Constitucion = 3
    

    'Extra

    'Party

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'
'Carga la lista de objetos
'
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))

'frmCargando.cargar.min = 0
'frmCargando.cargar.max = NumObjDatas
'frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = Val(Leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = Val(Leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = Val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = Val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = Val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
            'Pablo (ToxicWaste)
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
        Case eOBJType.otAnillo 'Pablo (ToxicWaste)
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            
    End Select
    
    ObjData(Object).Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
    ObjData(Object).RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).RazaDrow = Val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
    ObjData(Object).RazaElfa = Val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
    ObjData(Object).RazaGnoma = Val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    ObjData(Object).RazaHumana = Val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
    ObjData(Object).Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    
    'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
    Dim i As Integer
    Dim N As Integer
    Dim S As String
    For i = 1 To NUMCLASES
        S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
        N = 1
        Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
            N = N + 1
        Loop
        ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
    Next i
    
    ObjData(Object).DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    'frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Object

Set Leer = Nothing

Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer)

Dim LoopC As Long
Dim UserRaza As eRaza
Dim UserClase As eClass
UserClase = UserList(UserIndex).clase
For LoopC = 1 To 12
UserList(UserIndex).Invent.Object(LoopC).ObjIndex = 0
UserList(UserIndex).Invent.Object(LoopC).amount = 0
Next LoopC
UserList(UserIndex).Invent.Object(7).amount = 1
Select Case UserClase
    Case eClass.Mage
        UserRaza = eRaza.Humano
        UserList(UserIndex).Stats.MaxMAN = 2206
        UserList(UserIndex).Stats.MinMAN = 2206
        UserList(UserIndex).Stats.MaxHIT = 31
        UserList(UserIndex).Stats.MinHIT = 30
        UserList(UserIndex).Stats.MaxHP = 273
        UserList(UserIndex).Stats.MinHP = 273
        UserList(UserIndex).Invent.Object(2).ObjIndex = 986
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 660
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 662
        UserList(UserIndex).Invent.Object(4).amount = 1

    Case eClass.Druid, eClass.Bard
        UserRaza = eRaza.Elfo
        UserList(UserIndex).Stats.MaxMAN = 1610
        UserList(UserIndex).Stats.MinMAN = 1610
        UserList(UserIndex).Stats.MaxHIT = 60
        UserList(UserIndex).Stats.MinHIT = 59
        UserList(UserIndex).Stats.MaxHP = 312
        UserList(UserIndex).Stats.MinHP = 312
        UserList(UserIndex).Invent.Object(2).ObjIndex = 986
        UserList(UserIndex).Invent.Object(2).amount = 1
        If UserClase = eClass.Bard Then
            UserList(UserIndex).Invent.Object(11).ObjIndex = 399
            UserList(UserIndex).Invent.Object(11).amount = 1
            UserList(UserIndex).Invent.Object(4).ObjIndex = 404
            UserList(UserIndex).Invent.Object(4).amount = 1
            UserList(UserIndex).Invent.Object(3).ObjIndex = 132
            UserList(UserIndex).Invent.Object(3).amount = 1
            UserList(UserIndex).Invent.Object(5).ObjIndex = 696
            UserList(UserIndex).Invent.Object(5).amount = 1
            UserList(UserIndex).Invent.Object(12).ObjIndex = 365
            UserList(UserIndex).Invent.Object(12).amount = 1
        Else
            UserList(UserIndex).Invent.Object(3).ObjIndex = 365
            UserList(UserIndex).Invent.Object(3).amount = 1
            UserList(UserIndex).Invent.Object(4).ObjIndex = 208
            UserList(UserIndex).Invent.Object(4).amount = 1

        End If

    Case eClass.Cleric
        UserRaza = eRaza.Drow
        UserList(UserIndex).Stats.MaxMAN = 1610
        UserList(UserIndex).Stats.MinMAN = 1610
        UserList(UserIndex).Stats.MaxHIT = 70
        UserList(UserIndex).Stats.MinHIT = 69
        UserList(UserIndex).Stats.MaxHP = 312
        UserList(UserIndex).Stats.MinHP = 312
        UserList(UserIndex).Invent.Object(2).ObjIndex = 359
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 128
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 131
        UserList(UserIndex).Invent.Object(4).amount = 1
        UserList(UserIndex).Invent.Object(11).ObjIndex = 129
        UserList(UserIndex).Invent.Object(11).amount = 1
        UserList(UserIndex).Invent.Object(12).ObjIndex = 365
        UserList(UserIndex).Invent.Object(12).amount = 1
    Case eClass.Paladin
        UserRaza = eRaza.Humano
        UserList(UserIndex).Stats.MaxMAN = 702
        UserList(UserIndex).Stats.MinMAN = 702
        UserList(UserIndex).Stats.MaxHIT = 101
        UserList(UserIndex).Stats.MinHIT = 100
        UserList(UserIndex).Stats.MaxHP = 390
        UserList(UserIndex).Stats.MinHP = 390
        UserList(UserIndex).Invent.Object(2).ObjIndex = 195
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 128
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 131
        UserList(UserIndex).Invent.Object(4).amount = 1
        UserList(UserIndex).Invent.Object(11).ObjIndex = 129
        UserList(UserIndex).Invent.Object(11).amount = 1
        UserList(UserIndex).Invent.Object(12).ObjIndex = 365
        UserList(UserIndex).Invent.Object(12).amount = 1
    Case eClass.Assasin
        UserRaza = eRaza.Drow
        UserList(UserIndex).Stats.MaxMAN = 830
        UserList(UserIndex).Stats.MinMAN = 830
        UserList(UserIndex).Stats.MaxHIT = 101
        UserList(UserIndex).Stats.MinHIT = 100
        UserList(UserIndex).Stats.MaxHP = 312
        UserList(UserIndex).Stats.MinHP = 312
        UserList(UserIndex).Invent.Object(2).ObjIndex = 356
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 404
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 131
        UserList(UserIndex).Invent.Object(4).amount = 1
        UserList(UserIndex).Invent.Object(11).ObjIndex = 399
        UserList(UserIndex).Invent.Object(11).amount = 1
        UserList(UserIndex).Invent.Object(12).ObjIndex = 367
        UserList(UserIndex).Invent.Object(12).amount = 1
    Case eClass.Warrior
        UserRaza = eRaza.Enano
        UserList(UserIndex).Stats.MaxMAN = 0
        UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MaxHIT = 115
        UserList(UserIndex).Stats.MinHIT = 114
        UserList(UserIndex).Stats.MaxHP = 429
        UserList(UserIndex).Stats.MinHP = 429
        UserList(UserIndex).Invent.Object(2).ObjIndex = 243
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 128
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 131
        UserList(UserIndex).Invent.Object(4).amount = 1
        UserList(UserIndex).Invent.Object(7).ObjIndex = 479
        UserList(UserIndex).Invent.Object(7).amount = 1
        UserList(UserIndex).Invent.Object(8).ObjIndex = 480
        UserList(UserIndex).Invent.Object(8).amount = 10000
        UserList(UserIndex).Invent.Object(11).ObjIndex = 129
        UserList(UserIndex).Invent.Object(11).amount = 1
        UserList(UserIndex).Invent.Object(12).ObjIndex = 164 '625
        UserList(UserIndex).Invent.Object(12).amount = 1
    Case eClass.Hunter
        UserRaza = eRaza.Humano
        UserList(UserIndex).Stats.MaxMAN = 0
        UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MaxHIT = 70
        UserList(UserIndex).Stats.MinHIT = 60
        UserList(UserIndex).Stats.MaxHP = 390
        UserList(UserIndex).Stats.MinHP = 390
        UserList(UserIndex).Invent.Object(2).ObjIndex = 360
        UserList(UserIndex).Invent.Object(2).amount = 1
        UserList(UserIndex).Invent.Object(4).ObjIndex = 404
        UserList(UserIndex).Invent.Object(4).amount = 1
        UserList(UserIndex).Invent.Object(3).ObjIndex = 370
        UserList(UserIndex).Invent.Object(3).amount = 1
        UserList(UserIndex).Invent.Object(7).ObjIndex = 553
        UserList(UserIndex).Invent.Object(11).ObjIndex = 665
        UserList(UserIndex).Invent.Object(11).amount = 1
        UserList(UserIndex).Invent.Object(12).ObjIndex = 365
        UserList(UserIndex).Invent.Object(12).amount = 1
    Case Else
        UserClase = eClass.Cleric
        UserRaza = eRaza.Drow
        UserList(UserIndex).Stats.MaxMAN = 1610
        UserList(UserIndex).Stats.MinMAN = 1610
        UserList(UserIndex).Stats.MaxHIT = 70
        UserList(UserIndex).Stats.MinHIT = 65
        UserList(UserIndex).Stats.MaxHP = 312
        UserList(UserIndex).Stats.MinHP = 312
End Select
UserList(UserIndex).raza = UserRaza
UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 18 + ModRaza(UserRaza).Fuerza
UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 18 + ModRaza(UserRaza).Agilidad
UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = 18 + ModRaza(UserRaza).Inteligencia
UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = 18 + ModRaza(UserRaza).Carisma
UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = 18 + ModRaza(UserRaza).Constitucion

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = 100
Next LoopC

For LoopC = 1 To 10
  UserList(UserIndex).Invent.Object(LoopC).Equipped = 0
Next LoopC

UserList(UserIndex).Stats.GLD = 0
UserList(UserIndex).Stats.Banco = 0

UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100
UserList(UserIndex).Stats.MaxSta = 999
UserList(UserIndex).Stats.MinSta = 999
UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100

UserList(UserIndex).Stats.SkillPts = 0

UserList(UserIndex).Invent.Object(1).ObjIndex = 474
UserList(UserIndex).Invent.Object(1).amount = 1
UserList(UserIndex).Invent.BarcoSlot = 1
UserList(UserIndex).Invent.BarcoObjIndex = 474

UserList(UserIndex).Invent.Object(6).ObjIndex = 38
UserList(UserIndex).Invent.Object(6).amount = 1


If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
   UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
   UserClase = eClass.Assasin Or UserClase = eClass.Paladin Then
        UserList(UserIndex).Stats.UserHechizos(1) = 1
        UserList(UserIndex).Stats.UserHechizos(2) = 2
        UserList(UserIndex).Stats.UserHechizos(3) = 11
        UserList(UserIndex).Stats.UserHechizos(4) = 5
        UserList(UserIndex).Stats.UserHechizos(5) = 41
        UserList(UserIndex).Stats.UserHechizos(6) = 31
        UserList(UserIndex).Stats.UserHechizos(7) = 14
        UserList(UserIndex).Stats.UserHechizos(8) = 15
        UserList(UserIndex).Stats.UserHechizos(9) = 23
        UserList(UserIndex).Stats.UserHechizos(10) = 25
        UserList(UserIndex).Stats.UserHechizos(11) = 24
        UserList(UserIndex).Stats.UserHechizos(12) = 10
        UserList(UserIndex).Invent.Object(7).ObjIndex = 37
End If
If UserList(UserIndex).dios = True Then
        UserList(UserIndex).Stats.UserHechizos(1) = 32
        UserList(UserIndex).Stats.UserHechizos(2) = 34
End If
If UserClase = eClass.Druid Then
        UserList(UserIndex).Stats.UserHechizos(1) = 29
        UserList(UserIndex).Stats.UserHechizos(2) = 42
End If


Call DarCuerpoYCabeza(UserIndex)
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco


End Sub




Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
UserList(UserIndex).Reputacion.AsesinoRep = 1000
UserList(UserIndex).Reputacion.BandidoRep = 1000
UserList(UserIndex).Reputacion.BurguesRep = 1000
UserList(UserIndex).Reputacion.LadronesRep = 1000
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 1000
UserList(UserIndex).Reputacion.Promedio = 1000
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer)
'
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'
Dim LoopC As Long
Dim ln As String
UserList(UserIndex).bando = eNone
UserList(UserIndex).Faccion.ArmadaReal = 0
UserList(UserIndex).Faccion.FuerzasCaos = 0
UserList(UserIndex).Faccion.CiudadanosMatados = 0
UserList(UserIndex).Faccion.CriminalesMatados = 0
UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0
UserList(UserIndex).Faccion.RecibioExpInicialReal = 0
UserList(UserIndex).Faccion.RecompensasCaos = 0
UserList(UserIndex).Faccion.RecompensasReal = 0
UserList(UserIndex).Faccion.Reenlistadas = 0
UserList(UserIndex).Faccion.NivelIngreso = 0
UserList(UserIndex).Faccion.FechaIngreso = 0
UserList(UserIndex).Faccion.MatadosIngreso = 0
UserList(UserIndex).Faccion.NextRecompensa = 0

UserList(UserIndex).flags.Muerto = 1
UserList(UserIndex).flags.Escondido = 0

UserList(UserIndex).flags.Hambre = 0
UserList(UserIndex).flags.Sed = 0
UserList(UserIndex).flags.Desnudo = 0
UserList(UserIndex).flags.Navegando = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Paralizado = 0
UserList(UserIndex).email = "a@a.c"

UserList(UserIndex).genero = Hombre
UserList(UserIndex).clase = Pirat
UserList(UserIndex).raza = Drow
UserList(UserIndex).Hogar = 1
UserList(UserIndex).Char.heading = SOUTH

    UserList(UserIndex).Char.body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco


UserList(UserIndex).Pos.map = servermap


Dim xx As Byte
Dim yy As Byte
Dim salirfor As Boolean
salirfor = False
For xx = 9 To 90
    If salirfor = False Then
        For yy = 9 To 90
            If MapData(servermap, xx, yy).trigger = eTrigger.RESUCIU Or MapData(UserList(UserIndex).Pos.map, xx, yy).trigger = eTrigger.RESUPK Then

                    If MapData(UserList(UserIndex).Pos.map, xx, yy).trigger = eTrigger.RESUPK And LegalPos(UserList(UserIndex).Pos.map, xx, yy, False, True) = True Then
UserList(UserIndex).Pos.X = xx
UserList(UserIndex).Pos.Y = yy
salirfor = True
                        Exit For
                    End If

            End If
        Next yy
    Else
        Exit For
    End If
Next xx
If salirfor = False Then
UserList(UserIndex).Pos.X = 50
UserList(UserIndex).Pos.Y = 50
End If
UserList(UserIndex).Invent.NroItems = 0
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
szReturn = vbNullString
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub CargarBackUp()
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    
    Call InitAreas
    
    'frmCargando.cargar.min = 0
   ' frmCargando.cargar.max = NumMaps
    'frmCargando.cargar.value = 0
    
    MapPath = "\MAPAS\"
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & map
        Call CargarMapa(map, tFileName)
        
        'frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
      
    Dim contadoS As Long
    
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    
    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(map, X, Y).trigger = TempInt
                If MapData(map, X, Y).trigger = RESUCIU Or MapData(map, X, Y).trigger = RESUPK Then contadoS = contadoS + 1
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(map, X, Y).TileExit.map
                Get FreeFileInf, , MapData(map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(map, X, Y).NpcIndex
                
                'If MapData(map, X, Y).NpcIndex > 0 Then
                '    'If MapData(Map, X, Y).NpcIndex > 499 Then
                '    '    npcfile = DatPath & "NPCs-HOSTILES.dat"
                '    'Else
                '        npcfile = DatPath & "NPCs.dat"
                '    'End If'
'
 '                   'Si el npc debe hacer respawn en la pos
  '                  'original la guardamos
   '                 If Val(GetVar(npcfile, "NPC" & MapData(map, X, Y).NpcIndex, "PosOrig")) = 1 Then
    '                    MapData(map, X, Y).NpcIndex = OpenNPC(MapData(map, X, Y).NpcIndex)
     '                   Npclist(MapData(map, X, Y).NpcIndex).Orig.map = map
      '                  Npclist(MapData(map, X, Y).NpcIndex).Orig.X = X
       '                 Npclist(MapData(map, X, Y).NpcIndex).Orig.Y = Y
        '            Else
         '               MapData(map, X, Y).NpcIndex = OpenNPC(MapData(map, X, Y).NpcIndex)
          '          End If
           '
            '        Npclist(MapData(map, X, Y).NpcIndex).Pos.map = map
             '       Npclist(MapData(map, X, Y).NpcIndex).Pos.X = X
              '      Npclist(MapData(map, X, Y).NpcIndex).Pos.Y = Y
               '
                '    Call MakeNPCChar(True, 0, MapData(map, X, Y).NpcIndex, map, X, Y)
                'End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(map, X, Y).ObjInfo.ObjIndex
                Get FreeFileInf, , MapData(map, X, Y).ObjInfo.amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    Dim nmap As String
    nmap = GetVar(MAPFl & ".dat", "Mapa" & map, "Name")
    If Not Len(nmap) > 0 Then nmap = "Sin nombre"
    MapInfo(map).name = nmap
    frmMain.mapax.AddItem MapInfo(map).name
    frmMain.mapax.ListIndex = 0
    MapInfo(map).Music = GetVar(MAPFl & ".dat", "Mapa" & map, "MusicNum")
    MapInfo(map).StartPos.map = Val(ReadField(1, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).StartPos.X = Val(ReadField(2, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).StartPos.Y = Val(ReadField(3, GetVar(MAPFl & ".dat", "Mapa" & map, "StartPos"), Asc("-")))
    MapInfo(map).MagiaSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "MagiaSinEfecto"))
    MapInfo(map).InviSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "InviSinEfecto"))
    MapInfo(map).ResuSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "ResuSinEfecto"))
    MapInfo(map).NoEncriptarMP = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "NoEncriptarMP"))
    MapInfo(map).maxusersx = contadoS
    If Val(GetVar(MAPFl & ".dat", "Mapa" & map, "Pk")) = 0 Then
        MapInfo(map).Pk = True
    Else
        MapInfo(map).Pk = False
    End If
    
    
    MapInfo(map).Terreno = GetVar(MAPFl & ".dat", "Mapa" & map, "Terreno")
    MapInfo(map).Zona = GetVar(MAPFl & ".dat", "Mapa" & map, "Zona")
    MapInfo(map).Restringir = GetVar(MAPFl & ".dat", "Mapa" & map, "Restringir")
    MapInfo(map).BackUp = Val(GetVar(MAPFl & ".dat", "Mapa" & map, "BACKUP"))
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.Description)
End Sub

Sub LoadSini()

Dim Temporal As Long

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = False
'Misc
'#If SeguridadAlkon Then

'Call Security.SetServerIp(GetVar(IniPath & "Server.ini", "INIT", "ServerIp"))

'#End If


Puerto = 7666
HideMe = 0
AllowMultiLogins = Val(True)
IdleLimit = 5
'Lee la version correcta del cliente
ULTIMAVERSION = "0.1.2"

PuedeCrearPersonajes = 1
CamaraLenta = 0
ServerSoloGMs = 0

ArmaduraImperial1 = 370
ArmaduraImperial2 = 372
ArmaduraImperial3 = 492
TunicaMagoImperial = 517
TunicaMagoImperialEnanos = 549
ArmaduraCaos1 = 379
ArmaduraCaos2 = 523
ArmaduraCaos3 = 383
TunicaMagoCaos = 518
TunicaMagoCaosEnanos = 558

VestimentaImperialHumano = 675
VestimentaImperialEnano = 676
TunicaConspicuaHumano = 679
TunicaConspicuaEnano = 682
ArmaduraNobilisimaHumano = 629
ArmaduraNobilisimaEnano = 681
ArmaduraGranSacerdote = 680

VestimentaLegionHumano = 677
VestimentaLegionEnano = 678
TunicaLobregaHumano = 683
TunicaLobregaEnano = 685
TunicaEgregiaHumano = 634
TunicaEgregiaEnano = 686
SacerdoteDemoniaco = 684

servermap = 1

EnTesting = 0
EncriptarProtocolosCriticos = 1

'Start pos
StartPos.map = 1
StartPos.X = 50
StartPos.Y = 50

'Intervalos
SanaIntervaloSinDescansar = 1200
StaminaIntervaloSinDescansar = 2
SanaIntervaloDescansar = 100
StaminaIntervaloDescansar = 5
IntervaloSed = 10000
IntervaloHambre = 10000
IntervaloVeneno = 500
IntervaloParalizado = 500
IntervaloInvisible = 500
IntervaloFrio = 15
IntervaloWavFx = 190
IntervaloInvocacion = 1001
IntervaloParaConexion = 3000
IntervaloUserPuedeCastear = 1400
frmMain.TIMER_AI.Interval = 250
frmMain.npcataca.Interval = 2000
IntervaloUserPuedeTrabajar = 1200
IntervaloUserPuedeAtacar = 1200
'TODO : Agregar estos intervalos al form!!!
IntervaloMagiaGolpe = 1100
IntervaloGolpeMagia = 1100
MinutosWs = 180
If MinutosWs < 60 Then MinutosWs = 180
IntervaloCerrarConexion = 5
IntervaloUserPuedeUsar = 200
IntervaloFlechasCazadores = 1150
IntervaloOculto = 500
'&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&

'Ressurect pos
ResPos.map = 1
ResPos.X = 50
ResPos.Y = 50
  
recordusuarios = 20
  
'Max users
Temporal = 20
If maxusers = 0 Then
    maxusers = Temporal
    ReDim UserList(1 To maxusers) As User
End If



Nix.map = 1
Nix.X = 50
Nix.Y = 50

Ullathorpe.map = 1
Ullathorpe.X = 50
Ullathorpe.Y = 50

Banderbill.map = 1
Banderbill.X = 50
Banderbill.Y = 50

Lindos.map = 1
Lindos.X = 50
Lindos.Y = 50

Arghal.map = 1
Arghal.X = 50
Arghal.Y = 50




'Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'
'Escribe VAR en un archivo
'

writeprivateprofilestring Main, Var, value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)

'
'Author: Unknown
'Last modified: 23/01/2007
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'
Exit Sub
End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean
criminal = IIf(UserList(UserIndex).bando = eCUI, True, False)
End Function


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

End Sub
