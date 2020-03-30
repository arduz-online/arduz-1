VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColaArray"
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

Private m_maxelem As Long

Private m_array() As String
Private m_lastelem As Integer
Private m_firstelem As Integer
Private m_size As Integer

Public Function IsEmpty() As Boolean
IsEmpty = m_size = 0
End Function

Public Function IsFull() As Boolean
'IsFull = m_lastelem = m_maxelem
IsFull = m_size = m_maxelem
End Function

Public Function Push(ByVal aString As String) As Boolean

If Not Me.IsFull Then
    
    If Me.IsEmpty Then m_firstelem = 1
    
    m_lastelem = m_lastelem + 1
    If (m_lastelem > m_maxelem) Then m_lastelem = m_lastelem - m_maxelem
    m_size = m_size + 1
    m_array(m_lastelem) = aString
    
    Push = True
Else
    Push = False
End If


End Function

Public Function Pop() As String

If Not Me.IsEmpty Then
    
    Pop = m_array(m_firstelem)
    m_firstelem = m_firstelem + 1
    If (m_firstelem > m_maxelem) Then m_firstelem = m_firstelem - m_maxelem
    m_size = m_size - 1
    
    'If m_firstelem > m_lastelem And m_size = 0 Then
    If m_size = 0 Then
            m_lastelem = 0
            m_firstelem = 0
            m_size = 0
    End If
Else
    Pop = vbNullString
    
End If

End Function

Private Sub Class_Initialize()
m_lastelem = 0
m_firstelem = 0
m_size = 0
m_maxelem = 300

ReDim m_array(1 To m_maxelem)
End Sub

Public Property Get MaxElems() As Long
MaxElems = m_maxelem
End Property

Public Property Let MaxElems(ByVal lNewValue As Long)
m_maxelem = lNewValue
ReDim Preserve m_array(1 To m_maxelem)

End Property
