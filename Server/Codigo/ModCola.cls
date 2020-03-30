VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'

Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()
On Error Resume Next

Dim i As Integer
For i = 1 To Me.Longitud
    Cola.Remove FRENTE
Next i

End Sub

Public Property Get Longitud() As Integer
Longitud = Cola.Count
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
IndexValido = i >= 1 And i <= Me.Longitud
End Function

Private Sub Class_Initialize()
Set Cola = New Collection
End Sub

Public Function VerElemento(ByVal index As Integer) As String
On Error Resume Next
    If IndexValido(index) Then
        'Pablo
        VerElemento = UCase$(Cola.Item(index))
        '/Pablo
        'VerElemento = Cola(Index)
    Else
        VerElemento = 0
    End If
End Function


Public Sub Push(ByVal Nombre As String)
On Error Resume Next
'Mete elemento en la cola
'Pablo
Dim aux As String
aux = time$ + " " + UCase$(Nombre)
Call Cola.Add(aux)
'/Pablo

'Call Cola.Add(UCase$(Nombre))
End Sub

Public Function Pop() As String
On Error Resume Next
'Quita elemento de la cola
If Cola.Count > 0 Then
    Pop = Cola(FRENTE)
    Call Cola.Remove(FRENTE)
Else
    Pop = 0
End If
End Function

Public Function PopByVal() As String
On Error Resume Next
'Call LogTarea("PopByVal SOS")

'Quita elemento de la cola
If Cola.Count > 0 Then
    PopByVal = Cola.Item(1)
Else
    PopByVal = 0
End If

End Function

Public Function Existe(ByVal Nombre As String) As Boolean
On Error Resume Next
  
 Dim V As String
 Dim i As Integer
 Dim NombreEnMayusculas As String
 NombreEnMayusculas = UCase$(Nombre)
 
 For i = 1 To Me.Longitud
        'Pablo
        V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
        '/Pablo
        'V = Me.VerElemento(i)
        If V = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
 Next
 Existe = False

End Function

Public Sub Quitar(ByVal Nombre As String)
On Error Resume Next
Dim V As String
Dim i As Integer
Dim NombreEnMayusculas As String
 
NombreEnMayusculas = UCase$(Nombre)
 
For i = 1 To Me.Longitud
    'Pablo
    V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
    '/Pablo
    'V = Me.VerElemento(i)
    If V = NombreEnMayusculas Then
        Call Cola.Remove(i)
        Exit Sub
    End If
Next i
 
End Sub

Public Sub QuitarIndex(ByVal index As Integer)
On Error Resume Next
If IndexValido(index) Then Call Cola.Remove(index)
End Sub


Private Sub Class_Terminate()
'Destruimos el objeto Cola
Set Cola = Nothing
End Sub
