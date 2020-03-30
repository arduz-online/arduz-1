VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiMassClon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MaximoPersonajesPorIP = 15
Private m_coleccion As New Collection

Public Function MaxPersonajes(sIp As String) As Boolean
    Dim i As Long

    For i = 1 To m_coleccion.Count
        If m_coleccion.Item(i).ip = sIp Then
            m_coleccion.Item(i).PersonajesCreados = m_coleccion.Item(i).PersonajesCreados + 1
            MaxPersonajes = (m_coleccion.Item(i).PersonajesCreados > MaximoPersonajesPorIP)
            If MaxPersonajes Then m_coleccion.Item(i).PersonajesCreados = 16
            Exit Function
        End If
    Next i

#If SeguridadAlkon Then
    Dim oIp As New UserIpAdress
    oIp.ip = sIp
    oIp.PersonajesCreados = 1
    m_coleccion.Add oIp
#End If

    MaxPersonajes = False
    Exit Function
End Function

Public Function VaciarColeccion()

On Error GoTo Errhandler

Dim i As Integer

For i = 1 To m_coleccion.Count
   Call m_coleccion.Remove(1)
Next


Exit Function
Errhandler:
    Call LogError("Error en RestarConexion " & Err.description)
End Function


