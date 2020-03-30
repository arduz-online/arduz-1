VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public INT_ATTACK As Integer
Public INT_ARROWS As Integer
Public INT_CAST_SPELL As Integer
Public INT_CAST_ATTACK As Integer
Public INT_WORK As Integer
Public INT_USEITEMU As Integer
Public INT_USEITEMDCK As Integer
Public INT_SENTRPU As Integer


Private Sub Class_Initialize()
    INT_ATTACK = 1301
    INT_ARROWS = 1151
    INT_CAST_SPELL = 1051
    INT_CAST_ATTACK = 1151
    INT_WORK = 701
    INT_USEITEMU = 451
    INT_USEITEMDCK = 221
    INT_SENTRPU = 2001
End Sub
