Attribute VB_Name = "Mod_TCP"
Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
End Function

Sub Login()
    Call WriteLoginExistingChar
    DoEvents
    Call FlushBuffer
End Sub
