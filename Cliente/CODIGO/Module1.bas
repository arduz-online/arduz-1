Attribute VB_Name = "TAMAGOCHI"
Option Explicit

Type ePJ
    Nick As String
    clan As String
    desc As String
    frags As Long
    Puntos As Long
    muertes As Long
    ekipo As eKip
    id As Long
    gm As Boolean
    bot As Boolean
End Type

Public Enum eKip
    eNone = 0
    eCUI = 1
    ePK = 2
End Enum

Public pjs(40) As ePJ
Public Ekipos(3) As ekipo

Type ekipo
    Nombre As String
    color As Long
    num As Integer
    personajes(40) As Integer
End Type
Public totalxs As Integer

Public hamachi As Boolean

Public renderasd As Boolean
