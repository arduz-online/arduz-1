VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Arduz Online"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
If frmMain.Visible = True Then
frmMain.Visible = True
frmMain.Show
frmMain.SetFocus
frmMain.WindowState = vbNormal
Else
If frmConnect.Visible = True Then frmConnect.SetFocus
If frmOldPersonaje.Visible = True Then frmOldPersonaje.SetFocus
End If
End Sub

Private Sub Form_Load()
Unload Me
End Sub
