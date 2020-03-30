Attribute VB_Name = "dYERBAS"
Option Explicit
Public Const INVALID_INDEX As Integer = 0




Public Function Porcentaje(ByVal total As Long, ByVal Porc As Long) As Long
Porcentaje = (total * Porc) / 100
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.map - wp2.map) * 100)
End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double
Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
End Function

Sub Accion(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

End Sub



Public Function CharIndexToUserIndex(ByVal CharIndex As Integer) As Integer
    CharIndexToUserIndex = CharList(CharIndex)
    
    If CharIndexToUserIndex < 1 Or CharIndexToUserIndex > MaxUsers Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
    
    If UserList(CharIndexToUserIndex).Char.CharIndex <> CharIndex Then
        CharIndexToUserIndex = INVALID_INDEX
        Exit Function
    End If
End Function





#Const MODO_INVISIBILIDAD = 0

' cambia el estado de invisibilidad a 1 o 0 dependiendo del modo: true o false
'
Public Sub PonerInvisible(ByVal UserIndex As Integer, ByVal estado As Boolean)
#If MODO_INVISIBILIDAD = 0 Then

UserList(UserIndex).flags.invisible = IIf(estado, 1, 0)
UserList(UserIndex).flags.Oculto = IIf(estado, 1, 0)
UserList(UserIndex).Counters.Invisibilidad = 0

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, Not estado))


#Else

Dim EstadoActual As Boolean

' Está invisible ?
EstadoActual = (UserList(UserIndex).flags.invisible = 1)

'If EstadoActual <> Modo Then
    If Modo = True Then
        ' Cuando se hace INVISIBLE se les envia a los
        ' clientes un Borrar Char
        UserList(UserIndex).flags.invisible = 1
'        'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
        Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.map, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex))
    Else
        
    End If
'End If

#End If
End Sub


