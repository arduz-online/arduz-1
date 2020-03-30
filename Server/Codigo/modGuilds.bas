Attribute VB_Name = "modGuilds"
'
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
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

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






