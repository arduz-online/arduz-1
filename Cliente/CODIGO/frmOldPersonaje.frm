VERSION 5.00
Begin VB.Form frmOldPersonaje 
   BackColor       =   &H00374657&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1965
      Left            =   15
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   0
      Top             =   360
      Width           =   5130
      Begin VB.CheckBox sp 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Guardar contraseña"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox PasswordTxt 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox NameTxt 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Clave: (si no tenés, dejar en blanco)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Image1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingresar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Image1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Volver"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(si no estás registrado, usá cualquier nick, sin contraseña)"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFE96C&
      BackStyle       =   0  'Transparent
      Caption         =   "Ingresar al servidor..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   75
      Width           =   4815
   End
End
Attribute VB_Name = "frmOldPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Const textoKeypad = "Utilice el teclado como protección contra keyloggers. Seleccione el password con el mouse y presione <ENTER> al finalizar"
Private Const textoSeguir = "Conectarse al juego" & vbNewLine & "con el usuario y" & vbNewLine & "clave seleccionadas"
Private Const textoSalir = "Volver a la pantalla principal" & vbNewLine & "para crear personajes o recuperar" & vbNewLine & "contraseñas"

Private Sub Form_Load()
Dim j
For Each j In Image1()
    j.Tag = "0"
Next

frmOldPersonaje.NameTxt.Text = GetSetting(App.EXEName, "USER", "act", "Usuario")
If Len(NameTxt.Text) > 0 Then
'    If GetSetting(App.EXEName, "USERS", NameTxt.Text, "NOPEn") <> "NOPEn" Then
'    PasswordTxt.Text = GetSetting(App.EXEName, "USERS", NameTxt.Text, "NOPEn")
'    End If
verpasswD
End If

End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
       
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
#End If
        
        'update user info
        UserName = NameTxt.Text
        Dim aux As String
        aux = PasswordTxt.Text
        SaveSetting App.EXEName, "USER", "act", UserName
        If sp.Value = vbChecked Then Call SaveSetting(App.EXEName, "USERS", UserName, aux)
        If LenB(aux) < 1 Then aux = "NOTIENEPASSWD"
        If LenB(UserName) < 1 Then UserName = "Invitado"
#If SeguridadAlkon Then
        UserPassword = MD5.GetMD5String(aux)
        Call MD5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            EstadoLogin = Normal
            frmMain.pasarme
#If UsarWrench = 1 Then
            frmMain.Socket1.HostName = frmConnect.IPTxt.Text
            frmMain.Socket1.RemotePort = 7666 'frmConnect.PortTxt.Text
            frmMain.Socket1.Connect
#Else
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
        End If
        
    Case 1
        Me.Visible = False
End Select
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case index
    Case 0
        If Image1(0).Tag = "0" Then
            Image1(0).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
        End If
    Case 1
        If Image1(1).Tag = "0" Then
            Image1(1).Tag = "1"
            Call Audio.PlayWave(SND_OVER)
        End If
End Select
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(0)
    End If
End Sub

Private Sub NameTxt_LostFocus()
If Len(NameTxt.Text) = 0 Then
MsgBox "Ingrese un nick!"
NameTxt.SetFocus
End If
End Sub

Private Sub PasswordTxt_Click()
'If LenB(PasswordTxt) = 0 Then
verpasswD
End Sub

Public Sub verpasswD()
If Len(NameTxt.Text) = 0 Then NameTxt.Text = "Usuario"
On Error Resume Next
If GetSetting(App.EXEName, "USERS", NameTxt.Text, "NOPEn") <> "NOPEn" Then
PasswordTxt.Text = GetSetting(App.EXEName, "USERS", NameTxt.Text, "NOPEn")
End If
End Sub

Private Sub PasswordTxt_GotFocus()
If LenB(PasswordTxt) = 0 Then verpasswD
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(0)
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1(0).Tag = "0"
Image1(1).Tag = "0"
End Sub
