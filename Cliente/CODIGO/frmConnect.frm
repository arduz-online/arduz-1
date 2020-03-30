VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Arduz Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox hamaa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hamachi servers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   1770
   End
   Begin CLIENTE.ListadoServers lsts 
      Height          =   4215
      Left            =   6120
      TabIndex        =   3
      Top             =   3720
      Width           =   5415
      _ExtentX        =   6588
      _ExtentY        =   4471
      ColorSombra     =   8421504
      ColorLabel      =   7506330
      ColorDireccion  =   12632256
      ColorFondo      =   0
      BeginProperty TipoLetraLabels {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TipoLetraDireccion {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PunteroItems    =   0
      PunteroImagenItems=   "frmConnect.frx":59CE6
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   11040
      TabIndex        =   0
      Text            =   "7666"
      Top             =   3360
      Width           =   555
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   3360
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wssvr 
      Left            =   720
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image8 
      Height          =   1380
      Index           =   1
      Left            =   3120
      MouseIcon       =   "frmConnect.frx":59D02
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click para ir a la web."
      Top             =   720
      Width           =   5970
   End
   Begin VB.Image Image8 
      Height          =   300
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmConnect.frx":59E54
      MousePointer    =   99  'Custom
      Top             =   8400
      Width           =   3570
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   240
      MouseIcon       =   "frmConnect.frx":59FA6
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   1170
   End
   Begin VB.Image Image6 
      Height          =   420
      Left            =   240
      MouseIcon       =   "frmConnect.frx":5A0F8
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Image Image5 
      Height          =   420
      Left            =   240
      MouseIcon       =   "frmConnect.frx":5A24A
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   2730
   End
   Begin VB.Image Actualizar 
      Height          =   420
      Left            =   6120
      MouseIcon       =   "frmConnect.frx":5A39C
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   2490
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   240
      MouseIcon       =   "frmConnect.frx":5A4EE
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   3090
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   9240
      MouseIcon       =   "frmConnect.frx":5A640
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   240
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   3285
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitando lista de servidores..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   7920
      Width           =   3495
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Actualizar_Click()
lblStatus.Visible = True
lblStatus.Caption = "Solicicitando lista de servidores..."
Call Audio.PlayWave(SND_CLICK)
lsts.Resetear
On Error Resume Next
    With wssvr
        .Close
        .Protocol = sckUDPProtocol
        .RemoteHost = "255.255.255.255"
        .LocalPort = 4112
        .RemotePort = 4111
        .Bind 4112
    End With
    DoEvents
    wssvr.SendData CStr("IP")
Dim RD As String
If Inet1.StillExecuting = True Then Exit Sub
RD = Inet1.OpenURL("http://ao.noicoder.com/u.php?a=list&hamachi=" & CInt(hamaa.Value) & "&v=" & App.Major & "." & App.Minor & "." & App.Revision)
If RD Like "*NUEVAVERSION_*" Then
MsgBox "Hay una nueva version disponible en la web, para poder jugar nesecitás descargarla."
    If MsgBox("¿Querés descargar la nueva version o parche ahora?", vbYesNo) = vbYes Then
        Call ShellExecute(0, "Open", "http://ao.noicoder.com/?a=descargar&v=" & App.Major & "." & App.Minor & "." & App.Revision, "", App.path, 0)
    End If
End If
crearlista RD
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
'On Error Resume Next
Call Actualizar_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Make Server IP and Port box visible
'If KeyCode = vbKeyI And Shift = vbCtrlMask Then
'    'Port
'    PortTxt.Visible = True
'    'Label4.Visible = True
'    'Server IP
'    PortTxt.Text = "7666"
'    IPTxt.Text = "localhost"
'    IPTxt.Visible = True
'    'Label5.Visible = True
'    KeyCode = 0
'    Exit Sub
'End If

End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]
End Sub

Private Sub Image1_Click(index As Integer)
Call ShellExecute(0, "Open", "http://ao.noicoder.com/index.php?a=mi_cuenta", "", App.path, 0)
End Sub

Private Sub Image2_Click()
If Len(IPTxt.Text) > 6 And Len(PortTxt.Text) > 0 Then
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    DoEvents
    Call Audio.PlayWave(SND_CLICK)
    frmOldPersonaje.Show vbModal
End If
End Sub

Private Sub Image3_Click()
On Error GoTo errh
Call Shell(App.path & "\SERVER.EXE", vbNormalFocus)
Exit Sub
errh:
Call MsgBox("No se encuentra el programa Server.exe", vbCritical, "Arduz Online")
End Sub

Private Sub Image5_Click()
Call ShellExecute(0, "Open", "http://ao.noicoder.com/?a=ranking", "", App.path, 0)
Exit Sub
End Sub

Private Sub Image6_Click()
Call ShellExecute(0, "Open", "http://ao.noicoder.com/?a=ayuda", "", App.path, 0)
End Sub

Private Sub Image7_Click()
prgRun = False
End Sub



Private Sub Image8_Click(index As Integer)
Call ShellExecute(0, "Open", "http://ao.noicoder.com/", "", App.path, 0)
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
lblStatus.Visible = True
Select Case State

Case icResolvingHost
lblStatus.Visible = True
lblStatus.Caption = "Conectado al servidor principal"
Case icRequesting ' 5
lblStatus.Visible = True
lblStatus.Caption = "Solicitando lista"
Case icResponseReceived ' 8
lblStatus.Caption = "Lista recibida"
lblStatus.Visible = False
Case icDisconnecting ' 9
lblStatus.Caption = "Disconnecting"
Case icDisconnected ' 10
lblStatus.Caption = "Disconnected"
lblStatus.Visible = False
End Select
End Sub

Public Sub crearlista(RD As String)
On Error Resume Next
If RD = "" Then Exit Sub
If InStr(RD, "@") <= 0 Then Exit Sub
Dim i As Integer
Dim j As Integer
Dim parts() As String
Dim parts1() As String
  parts = Split(RD, "@|")
  j = UBound(parts)
  For i = 1 To UBound(parts)
    parts1 = Split(parts(i), "ç")
    If UBound(parts1) > 0 Then
    Dim asd As Object
    addsvr parts1(2), parts1(3), parts1(4), parts1(1), parts1(0), parts1(1)
    End If
  Next
End Sub


Sub addsvr(nombresv As String, mapasv As String, jugadoressv As String, modosv As String, ipsv As String, portsv As String)
On Error Resume Next
'Dim pin As Integer
'pin = PingIp(ipsv, 115, 0)
'Debug.Print pin
lsts.AddItem nombresv, ipsv, portsv, mapasv, "-1", jugadoressv
End Sub

Private Sub lsts_Click(index As Integer, Item As String, direccion As String, Puerto As Long)
Call Audio.PlayWave(SND_CLICK)
PortTxt.Text = Puerto
IPTxt.Text = direccion
End Sub

Private Sub lsts_DblClick(index As Integer, Item As String, direccion As String, Puerto As Long)
DoEvents
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
frmOldPersonaje.Show vbModal
End Sub

Private Sub version_Click()

End Sub

Private Sub wssvr_DataArrival(ByVal bytesTotal As Long)
    Dim msg As String
    wssvr.GetData msg, vbString
    If msg Like "*@|*" Then
        Dim parts1() As String
        parts1 = Split(msg, "ç")
        If UBound(parts1) > 0 Then
        addsvr parts1(3), parts1(4), parts1(5), parts1(2), wssvr.RemoteHostIP, parts1(2)
        End If
    End If
End Sub
