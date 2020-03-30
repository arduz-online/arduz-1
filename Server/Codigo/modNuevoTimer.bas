Attribute VB_Name = "modNuevoTimer"
Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
    If Actualizar Then
        UserList(UserIndex).Counters.TimerLanzarSpell = TActual
    End If
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
    If Actualizar Then
        UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
    End If
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    
    If UserList(UserIndex).Counters.TimerMagiaGolpe > UserList(UserIndex).Counters.TimerLanzarSpell Then
        Exit Function
    End If
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerMagiaGolpe = TActual
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
        End If
        IntervaloPermiteMagiaGolpe = True
    Else
        IntervaloPermiteMagiaGolpe = False
    End If
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    
    If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
        Exit Function
    End If
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
    IntervaloPermiteTrabajar = True
Else
    IntervaloPermiteTrabajar = False
End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
    If Actualizar Then UserList(UserIndex).Counters.TimerUsar = TActual
    IntervaloPermiteUsar = True
Else
    IntervaloPermiteUsar = False
End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
    IntervaloPermiteUsarArcos = True
Else
    IntervaloPermiteUsarArcos = False
End If

End Function


