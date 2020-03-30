VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arduz Server"
   ClientHeight    =   5760
   ClientLeft      =   1950
   ClientTop       =   1740
   ClientWidth     =   4905
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame3 
      Caption         =   "Servidor"
      Height          =   2055
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   4695
      Begin VB.ComboBox mapax 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox svrname 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox hamaa 
         Alignment       =   1  'Right Justify
         Caption         =   "Hamachi"
         Height          =   255
         Left            =   3360
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Porttt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Text            =   "7666"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox ronda 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tiempo limite:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   270
         Width           =   1495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   1030
         TabIndex        =   32
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   200
         TabIndex        =   31
         Top             =   1000
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Puerto: "
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1330
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Bots"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4050
      Width           =   669
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label asddsadsa 
         Alignment       =   2  'Center
         Caption         =   "Iniciando servidor..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   30
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Restart"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSWinsockLib.Winsock wssvr 
      Left            =   3720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Minimizar"
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer TimerControl 
      Interval        =   3000
      Left            =   3720
      Top             =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bots"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4065
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ComboBox mankoo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":1042
         Left            =   1680
         List            =   "frmMain.frx":1044
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Matar todos"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   " Agregar Rojo"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   " Agregar Azul"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Mankism 
         Alignment       =   1  'Right Justify
         Caption         =   "Nivel de los bots:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   1455
      End
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   3720
      Top             =   480
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer AutoSave 
      Interval        =   60000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   120
   End
   Begin VB.CheckBox envrank 
      Caption         =   "Enviar datos al ranking"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Iniciarsv 
      Caption         =   "Iniciar Servidor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Admin"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command8 
         Caption         =   "BAN"
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox invii 
         Caption         =   "Invisbilidad"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox estuu 
         Caption         =   "Estupidez"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox resuu 
         Caption         =   "Resucitar"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox ffire 
         Caption         =   "Friendly Fire"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox deathms 
         Caption         =   "Deathmatch"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox fatu 
         Caption         =   "Invocaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox adminpas 
         Height          =   315
         Left            =   2400
         TabIndex        =   18
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton C123 
         Caption         =   "ECHAR"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ADMIN"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboPjs 
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Numero de usuarios: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   41
         Top             =   0
         Width           =   1875
      End
      Begin VB.Label Escuch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Contrase�a de admin:"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   33
      Top             =   3600
      Width           =   4365
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Men�"
      Begin VB.Menu mnuSystray 
         Caption         =   "Esconder"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uid As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205


Private Enum en_OptTypeEnable
    enProggy
    enPort
End Enum

'---- file open in api a.k.a. common dialog api
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private PROGSCOPE As NET_FW_SCOPE_
Private PORTSCOPE As NET_FW_SCOPE_
Private Protocol  As NET_FW_IP_PROTOCOL_
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, id As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uid = id
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To maxusers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                Call Cerrar_Usuario(iUserIndex)
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaStatsES

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()

On Error GoTo Errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long
Static bool As Boolean
Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        Horas = 0
        
    End If
    MinsRunning = 0
End If
WEBCLASS.TryRequest
bool = Not bool
If bool = True Then
Call WEBCLASS.PingToWeb
End If

Minutos = Minutos + 1

Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
Errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)
    Resume Next
End Sub







Private Sub Command1_Click()
Call Command5_Click
Call Command7_Click
End Sub





Private Sub Check2_Click()
Frame2.Visible = Check2.value
botsact = Check2.value
If botsact = False Then Call Command4_Click
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To maxusers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

'Private Sub Command1_Click()
'Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
'End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
    Me.WindowState = vbNormal
End If

End Sub
'
'Private Sub Command2_Click()
'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
'End Sub

Private Sub Command10_Click()
Dim wp As WorldPos
wp.X = ALCOBA1_X + CInt(Rnd * 5)
wp.Y = ALCOBA1_Y + CInt(Rnd * 5)
Call CrearNPC(PRKING_NPC, servermap, wp)
End Sub

Private Sub Command11_Click()
'CrearClanPretoriano 50
Dim wp As WorldPos
If RandomNumber(0, 2) = 1 Then
    Call CrearNPC(PRCLER_NPC, servermap, wp, eKip.epk)
Else
    Call CrearNPC(PRMAGO_NPC, servermap, wp, eKip.epk)
End If
End Sub



Private Sub Command2_Click()
Dim wp As WorldPos
If RandomNumber(0, 2) = 1 Then
    Call CrearNPC(PRCLER_NPC, servermap, wp, eKip.eCUI)
Else
    Call CrearNPC(PRMAGO_NPC, servermap, wp, eKip.eCUI)
End If
End Sub

Private Sub Command3_Click()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
    If MsgBox("�Seguro quer�s hacer admin a " & cboPjs.Text & "?", vbYesNo) = vbYes Then
        UserList(tIndex).admin = Not UserList(tIndex).admin
    End If
End If

End Sub

Private Sub Command4_Click()
pretorianosVivos = 0
Dim i As Integer
        For i = 1 To 100
        If Npclist(i).flags.NPCActive = True Then Call QuitarNPC(i)
        Next i
End Sub

Private Sub Command5_Click()
If servermap <> mapax.ListIndex + 1 Then
    servermap = mapax.ListIndex + 1
    Call cambiarmapa
End If

mankismo = mankoo.ListIndex
svname = IIf(Len(svrname.Text) > 1, svrname.Text, "Nombre del Servidor")
svrname.Text = svname
rondaact = ronda.Enabled
rondaa = (ronda.ListIndex * 60 * 5)
If rondaa = 0 Then rondaa = 60
valeinvi = invii.value
valeestu = estuu.value
valeresu = resuu.value
adminpasswd = IIf(Len(svrname.Text) > 2, adminpas.Text, "CONTRASE�A")
If adminpasswd = "CONTRASE�A" Then
txStatus.Caption = "INGRESE UNA CONTRASE�A!!"
Else
txStatus.Caption = ""
End If
adminpas.Text = adminpasswd
SaveSetting App.EXEName, "SERVER", "NAME", svrname.Text
SaveSetting App.EXEName, "SERVER", "PASS", adminpasswd
enviarank = envrank.value
atacaequipo = ffire.value
fatuos = fatu.value
deathm = deathms.value
If enviarank = True Then
    WEBCLASS.PingToWeb
End If
If serverrunning Then enviaser
End Sub

Private Sub Command7_Click()
Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False
End Sub


Private Sub Check1_Click()
ronda.Enabled = Check1.value
End Sub

Private Sub deathms_Click()
If deathms.value = vbChecked Then
ffire.value = vbChecked
ffire.Enabled = False
Else
ffire.Enabled = True
End If
End Sub

Private Sub Form_Load()
Init_Hamachi
Me.Caption = "Arduz Server - "
'sckBroadcast.RemotePort = 1414 'Set the 'Remote listen port'
'sckBroadcast.AddressFamily = AF_INET
'sckBroadcast.Protocol = IPPROTO_UDP 'Use the UDP Protocol (MUST)
'sckBroadcast.SocketType = SOCK_DGRAM
'sckBroadcast.Broadcast = True 'Enable the broadcasting feature :)
'sckBroadcast.Binary = False 'Disable the sendpacket type to binary
'sckBroadcast.Blocking = False
'sckBroadcast.Action = SOCKET_OPEN 'Everything is set, enable/open the socket!
ronda.AddItem "1 Minuto"
ronda.AddItem "5 Minutos"
ronda.AddItem "10 Minutos"
ronda.AddItem "15 Minutos"
ronda.AddItem "20 Minutos"
ronda.AddItem "25 Minutos"
mankoo.AddItem "Experto"
mankoo.AddItem "Bueno"
mankoo.AddItem "Normal"
mankoo.AddItem "Semi Normal"
mankoo.AddItem "Semi casi manco"
mankoo.AddItem "Casi manko"
mankoo.AddItem "Semi manko"
mankoo.AddItem "Algo es Algo :P"
mankoo.ListIndex = 2
ronda.ListIndex = 0
svrname.Text = GetSetting(App.EXEName, "SERVER", "NAME")
adminpas.Text = GetSetting(App.EXEName, "SERVER", "PASS")
svname = svrname.Text
End Sub


Private Sub C123_Click()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
If UserList(tIndex).dios = True Then
WriteConsoleMsg tIndex, "TE QUIEREN HECHAR DEL SERVER, DESLOGUEA.", FontTypeNames.FONTTYPE_WARNING
Exit Sub
End If
    If MsgBox("��Seguro quer�s echar a " & cboPjs.Text & " del servidor?!", vbYesNo) = vbYes Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(tIndex).name & " ha sido hechado. ", FontTypeNames.FONTTYPE_SERVER))
        Call CloseSocket(tIndex)
    End If
End If

End Sub

Public Sub ActualizaListaPjs()
Dim LoopC As Long
With cboPjs
    .Clear
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
                .AddItem UserList(LoopC).name
                .ItemData(.NewIndex) = LoopC
        End If
    Next LoopC
End With
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To maxusers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To maxusers 'LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '�User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    .NumeroPaquetesPorMiliSec = 0
                    Call DoTileEvents(iUserIndex, .Pos.map, .Pos.X, .Pos.Y)
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    If .flags.Muerto = 0 Then
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo Then
                                If Not Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                    'No esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                    Else
                                    'esta descansando
                                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        'termina de descansar automaticamente
                                        If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False
                                        End If
                                        
                                    End If
                                End If
                            Else
                                If Not .flags.Descansar Then
                                'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                'esta descansando
                                    
                                    'Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    'termina de descansar automaticamente
                                    If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False
                                    End If
                                    
                                End If
                            End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim d_Chunk As Variant
Dim Datos As String
    If State = inetctlsobjects.StateConstants.icResponseCompleted Then
        d_Chunk = Inet1.GetChunk(1024, icString)
        Datos = Datos & d_Chunk
        Do
            DoEvents
            d_Chunk = Inet1.GetChunk(1024, icString)
            If Len(d_Chunk) = 0 Then
               Exit Do
            Else
              Datos = Datos & d_Chunk
            End If
        Loop
        WEBCLASS.PharseResultWeb Datos
        WEBCLASS.InetState = False
    ElseIf State = inetctlsobjects.StateConstants.icDisconnected Then
        WEBCLASS.InetState = False
    End If
End Sub

Private Sub Iniciarsv_Click()
If Len(svrname) < 4 Then
    txStatus.Caption = "Ingrese un nombre"
    svrname.SetFocus
    Exit Sub
End If
Porttt.Enabled = False
Call Porttt_Change
Call Porttt_LostFocus
Puerto = CLng(Porttt.Text)
hamaa.Enabled = False
envrank.Enabled = False
Picture1.Visible = True
DoEvents
botsact = False
Call Command4_Click
    txStatus.Caption = ""
    With wssvr
        .Close
        .Protocol = sckUDPProtocol
        .RemoteHost = "255.255.255.255"
        .LocalPort = 4111
        .RemotePort = 4112
        
        ' start listening for UDP packets
        .bind 4111
    End With

Call Command5_Click
Command6.Visible = True
Command5.Visible = True
Command1.Visible = True
servermap = mapax.ListIndex + 1
Call CrearClanPretoriano(50)
Command6.Enabled = True
Iniciarsv.Enabled = False
serverrunning = True
Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To maxusers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

With frmMain
    .AutoSave.Enabled = True
    .GameTimer.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If
Call WEBCLASS.CrearServerWeb
Picture1.Visible = False
End Sub



Private Sub mnuCerrar_Click()
Me.Hide
If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. �Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    If Iniciarsv.Enabled = False Then LimpiaWsApi
    WEBCLASS.BorrarServerWeb
    
    Dim f
    For Each f In Forms
        Unload f
        
    Next
    
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub



Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()
WEBCLASS.TryRequest
On Error Resume Next
Dim npc As Integer
For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc
End Sub

Private Sub packetResend_Timer()
'
'
'04/01/07
'Attempts to resend to the user all data that may be enqueued.
'
On Error GoTo Errhandler:
    Dim i As Long
    
    For i = 1 To maxusers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

Errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.Description)
    Resume Next
End Sub




Private Sub Porttt_Change()
If Not IsNumeric(Porttt.Text) Then Porttt.Text = "7666"

End Sub


Private Sub Porttt_KeyPress(KeyAscii As Integer)
Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ( _
        (ch >= "0" And ch <= "9") _
    ) And KeyAscii <> vbKeyBack Then
        ' Cancel the character.
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub Porttt_LostFocus()
If CLng(Porttt.Text) < 82 Or CLng(Porttt.Text) > 9999 Then Porttt.Text = "7666"
End Sub

Private Sub TIMER_AI_Timer()
If NumUsers = 0 Then Exit Sub
If botsact = False Then Exit Sub
'On Error GoTo ErrorHandler
Dim NpcIndex As Long
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado Then
                Call EfectoParalisisNpc(NpcIndex)
            End If
                e_p = esPretoriano(NpcIndex)
                If e_p > 0 And Npclist(NpcIndex).inerte = False Then
                    Select Case e_p
                        Case 1  ''clerigo
                            Call PRCLER_AI(NpcIndex)
                        Case 2  ''mago
                            Call PRMAGO_AI(NpcIndex)
                        Case 3  ''cazador
                            Call PRCAZA_AI(NpcIndex)
                        Case 4  ''rey
                            Call PRREY_AI(NpcIndex)
                        Case 5  ''guerre
                            Call PRGUER_AI(NpcIndex)
                    End Select
                Else
            If Npclist(NpcIndex).flags.Paralizado = 1 Then
                Call EfectoParalisisNpc(NpcIndex)
            Else                    'Usamos AI si hay algun user en el mapa
                    If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                       Call EfectoParalisisNpc(NpcIndex)
                    End If
                    
                    mapa = Npclist(NpcIndex).Pos.map
                    
                    If mapa > 0 Then
                        If MapInfo(mapa).NumUsers > 0 Then
                            If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                Call NPCAI(NpcIndex)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next NpcIndex
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.map)
    Call MuereNpc(NpcIndex, 0)
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub TimerControl_Timer()
bbmanda
End Sub

Sub bbmanda()
If NumUsers = 0 Then Exit Sub
Static conteoz As Integer
If deathm = True Then
    dLlevarRand
    conteoz = conteoz + 1
    If conteoz > 30 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor>Enviando ranking...", FontTypeNames.FONTTYPE_SERVER))
        WEBCLASS.enviarpjs
        conteoz = 0
    End If
Else
    roundstart
    If rondaact = True Then
        Dim i As Integer, k As Integer
        rondax = rondax + 1
        If rondax >= rondaa Then
            volverbases
            rondax = 0
        ElseIf rondax >= rondaa - 6 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("TERMINA EN " & rondaa - rondax, FontTypeNames.FONTTYPE_FIGHT))
        ElseIf rondax < 5 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("" & IIf((4 - rondax) <= 0, "Conteo>YA!", "Conteo>" & (4 - rondax)), FontTypeNames.FONTTYPE_GM))
        End If
    End If
End If
End Sub

Sub roundstart()
Dim i As Integer, k As Integer
Dim hacer As Boolean
Dim hacer1 As Boolean
Dim vivopk As Boolean
Dim vivociu As Boolean
Dim numpk, numciu As Integer

    numpk = UserBando(eKip.epk)
    numciu = UserBando(eKip.eCUI)
    vivopk = UserVivos(eKip.epk) > 0
    vivociu = UserVivos(eKip.eCUI) > 0
    
    
        If 
        (vivopk = False And numpk > 0) Or 
        (vivociu = False And numciu > 0) Or (
          (vivopk = False And numpk > 0)
          And
          (vivociu = False And numciu > 0)
          ) Then
            
            
            If vivociu = True Then
                winciu = winciu + 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("�EL EQUIPO ROJO GAN� LA RONDA!.", FontTypeNames.FONTTYPE_TALK))
                
            ElseIf vivopk = True Then
                winpk = winpk + 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("�EL EQUIPO AZUL GAN� LA RONDA!.", FontTypeNames.FONTTYPE_TALK))
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Puntajes:" & vbNewLine & "EQUIPO AZUL:" & winpk & vbNewLine & "EQUIPO ROJO:" & winciu, FontTypeNames.FONTTYPE_VENENO))
            For i = 1 To 100
                If Npclist(i).flags.NPCActive = True Then Call QuitarNPC(i)
            Next i
            rondax = 0
            
            If enviarank = True Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor>Enviando ranking...", FontTypeNames.FONTTYPE_SERVER))
                WEBCLASS.enviarpjs
            End If
            
            For i = 1 To maxusers 'LastUser
                    With UserList(i)
                           'Conexion activa?
                            If .ConnID <> -1 Then
                                '�User valido?
                                If .ConnIDValida And .flags.UserLogged Then
                                'Call UserDie(i)
                                    If .bando = epk And vivopk = True Then
                                        .Stats.puntos = .Stats.puntos + 100
                                        .Stats.puntosenv = .Stats.puntosenv + 100
                                    ElseIf .bando = eCUI And vivociu = True Then
                                        .Stats.puntos = .Stats.puntos + 100
                                        .Stats.puntosenv = .Stats.puntosenv + 100
                                    End If
                                    
                                    Call volverbase(i)

                                End If
                            End If
                    End With
                Next i
            


            
            Call CrearClanPretoriano(50)
            Exit Sub
        End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal id As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= maxusers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato id, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(id)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(id))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = id
        UserList(NewIndex).ip = TCPServ.GetIP(id)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(id) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal id As Long, ByVal MiDato As Long)
On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & id & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal id As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & id & " error:" & Err.Description)

End Sub



Private Sub tLluviaEvent_Timer()

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub wssvr_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
Private Sub wssvr_DataArrival(ByVal bytesTotal As Long)
    Dim msg As String
    On Error Resume Next
    ' Received message from client
    wssvr.GetData msg, vbString
    Debug.Print msg
    ' Check if message is from a "friendly" application (our client application)
    If msg Like "*IP*" Then
    Debug.Print msg & "asd"
        ' Broadcast back our IP and TCP port number
        wssvr.SendData "@|�" & wssvr.LocalIP & "�" & Puerto & "�" & svname & "�" & mapax.list(servermap - 1) & "�" & NumUsers
    End If
End Sub

