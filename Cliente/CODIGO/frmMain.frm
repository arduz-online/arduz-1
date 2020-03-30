VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":16CA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   4200
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer soycheater 
      Interval        =   894
      Left            =   5160
      Top             =   3000
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   195
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1980
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   7080
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   12000
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   16
      Top             =   8520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   6600
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   6120
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   5640
      Top             =   2520
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5160
      Top             =   2520
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8145
      Left            =   8280
      Picture         =   "frmMain.frx":5E901
      ScaleHeight     =   543
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   1
      Top             =   480
      Width           =   3585
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "EQUIPO ROJO"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EQUIPO AZUL"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4680
         Width           =   1335
      End
      Begin CLIENTE.GradientProgressBar vids 
         Height          =   375
         Left            =   390
         TabIndex        =   19
         Top             =   7050
         Width           =   2775
         _extentx        =   2355
         _extenty        =   450
         gradienttype    =   6
      End
      Begin CLIENTE.GradientProgressBar mans 
         Height          =   375
         Left            =   390
         TabIndex        =   18
         Top             =   6165
         Width           =   2775
         _extentx        =   2355
         _extenty        =   450
         gradienttype    =   10
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   540
         MouseIcon       =   "frmMain.frx":7C7EC
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   540
         MouseIcon       =   "frmMain.frx":7C93E
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   555
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   7
         Top             =   2280
         Width           =   2400
      End
      Begin VB.ListBox hlst 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2760
         Left            =   420
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.ListBox lstClases 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFC0&
         Height          =   2175
         Left            =   480
         TabIndex        =   20
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label cmdMoverHechi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label cmdMoverHechi 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPorcLvl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "33.33%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   2880
         TabIndex        =   15
         Top             =   6720
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Image cmdInfo 
         Height          =   405
         Left            =   2310
         MouseIcon       =   "frmMain.frx":7CA90
         MousePointer    =   99  'Custom
         Top             =   4830
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image CmdLanzar 
         Height          =   645
         Left            =   360
         MouseIcon       =   "frmMain.frx":7CBE2
         MousePointer    =   99  'Custom
         Top             =   4800
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label exp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   2
         Left            =   3000
         Top             =   6720
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   1
         Left            =   2055
         Top             =   6255
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   0
         Left            =   3000
         Top             =   6720
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   6000
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmMain.frx":7CD34
         MousePointer    =   99  'Custom
         Top             =   7680
         Width           =   3090
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   1
         Left            =   2040
         MouseIcon       =   "frmMain.frx":7CE86
         MousePointer    =   99  'Custom
         Top             =   5760
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   360
         Top             =   7560
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   360
         Top             =   7440
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H00003135&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00004040&
         Height          =   15
         Left            =   360
         Top             =   5760
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Personaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0072899A&
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   3105
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1800
         MouseIcon       =   "frmMain.frx":7CFD8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1290
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   240
         MouseIcon       =   "frmMain.frx":7D12A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1305
         Width           =   1605
      End
      Begin VB.Image InvEqu 
         Height          =   4395
         Left            =   165
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   4
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2760
         TabIndex        =   3
         Top             =   6720
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2280
         TabIndex        =   2
         Top             =   6720
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   195
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":7D27C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmMain.frx":7D2F9
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   4
      Left            =   9360
      MouseIcon       =   "frmMain.frx":7D44B
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   1290
   End
   Begin VB.Label FPS 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EBAC6&
      Height          =   255
      Left            =   5130
      TabIndex        =   24
      Top             =   60
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   10920
      MouseIcon       =   "frmMain.frx":7D59D
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   930
   End
   Begin VB.Image PicResu 
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   12120
      Picture         =   "frmMain.frx":7D6EF
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Coord 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,00,00)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   12120
      TabIndex        =   14
      Top             =   8730
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000000&
      Height          =   6240
      Left            =   105
      Top             =   2325
      Width           =   8190
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   5100
      TabIndex        =   25
      Top             =   45
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5100
      TabIndex        =   26
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   27
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   28
      Top             =   45
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   5130
      TabIndex        =   29
      Top             =   90
      Width           =   975
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
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

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public pri As Boolean

Public IsPlaying As Byte

Dim PuedeMacrear As Boolean



Sub pasarme()
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    picInv.Visible = False
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    InvEqu.Picture = LoadPicture()
    picInv.Visible = False
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    Label4.Enabled = False
    Label7.Enabled = False
    Command1.Visible = True
    Command2.Visible = True
    lstClases.Visible = True
    lstClases.Clear
    lstClases.AddItem "MAGO"
    lstClases.AddItem "CLERIGO"
    lstClases.AddItem "PALADIN"
    lstClases.AddItem "GUERRERO"
    lstClases.AddItem "CAZADOR"
    lstClases.AddItem "ASESINO"
    lstClases.AddItem "BARDO"
    lstClases.AddItem "DRUIDA"
    lstClases.listIndex = 0
End Sub



Private Sub cmdMoverHechi_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)
    If hlst.listIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case index
        Case 1 'subir
            If hlst.listIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(index, hlst.listIndex + 1)
    
    Select Case index
        Case 1 'subir
            sTemp = hlst.List(hlst.listIndex - 1)
            hlst.List(hlst.listIndex - 1) = hlst.List(hlst.listIndex)
            hlst.List(hlst.listIndex) = sTemp
            hlst.listIndex = hlst.listIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.listIndex + 1)
            hlst.List(hlst.listIndex + 1) = hlst.List(hlst.listIndex)
            hlst.List(hlst.listIndex) = sTemp
            hlst.listIndex = hlst.listIndex + 1
    End Select
End Sub

Private Sub Command1_Click()
Call Audio.PlayWave(SND_CLICK)
WriteEditChar "yo", eo_Class, lstClases.List(lstClases.listIndex), 1
Call Label4_Click
End Sub
Private Sub Command2_Click()
Call Audio.PlayWave(SND_CLICK)
WriteEditChar "yo", eo_Class, lstClases.List(lstClases.listIndex), 0
Call Label4_Click
End Sub



Private Sub Command3_Click()
    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario.jpg")
Call Label4_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
IScombate = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
    
    If (Not SendTxt.Visible) Then
            Select Case KeyCode
                Case vbKeyMultiply
                    Call WriteInvisible
                Case vbKeyM
                    Audio.MusicActivated = Not Audio.MusicActivated
                Case vbKeyA
                    Call AgarrarItem
                
                Case vbKeyE
                    Call EquiparItem
                
                Case vbKeyN
                    Nombres = Not Nombres

                Case vbKeyO
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                Case vbKeySpace
                    
                    IScombate = False
                Case vbKeyU
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case vbKeyL
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case vbKeyEnd
                        If UserMinMAN = UserMaxMAN Then Exit Sub
                            Call WriteMeditate
            End Select

        End If
    
    Select Case KeyCode
        Case vbKeyF4
            Call WriteQuit
            
        Case vbKeyControl
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            Call WriteAttack
        
        Case vbKeyReturn
                SendTxt.Visible = True
                IScombate = False
                SendTxt.SetFocus
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub



Private Sub FPS_Click()
Me.WindowState = vbNormal
renderasd = True
Me.Visible = True
EngineRun = True
End Sub



Private Sub Label1_Click()

End Sub

Private Sub lstClases_Click()
'Me.WindowState = vbNormal
'renderasd = True
'Me.Visible = True
'EngineRun = True
End Sub

Private Sub Macro_Timer()
PuedeMacrear = True

End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub


Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub



Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub soycheater_Timer()
Static ultima As Long
Static actual As Long
    If frmMain.WindowState <> vbMinimized And renderasd = True Then
        If pri = True Then
            actual = GetTickCount()
            If (actual - ultima - 210) > soycheater.Interval Then WriteBankStart 'MsgBox "Soy cheater."
        End If
        pri = True
        ultima = GetTickCount()
        MainTimer.ChechCheat
    Else
        ultima = GetTickCount()
        pri = False
    End If
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
Static dada As Boolean
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
    
If IScombate = True Then
        dada = Not dada
    Call WritePing
        If dada = True Then Call WriteRequestUserList
    
End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    Call WritePickUp
End Sub

Private Sub UsarItem()
If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Public Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''
Private Sub cmdLanzar_Click()
    If MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.listIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.listIndex <> -1 Then
        Call WriteSpellInfo(hlst.listIndex + 1)
    End If
End Sub

Private Sub DespInv_Click(index As Integer)
    Inventario.ScrollInventory (index = 0)
End Sub

Private Sub Form_Click()
Me.WindowState = vbNormal
renderasd = True
Me.Visible = True
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                

                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichTextBox(frmMain.RecTxt, "No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    'Static asdds As Integer
    'asdds = asdds + 1
    'If Not frmForo.Visible And Not frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    'End If
    'charlist(1).Head = HeadData(asdds)
    'Me.Caption = asdds
    'If asdds > 25 Then asdds = -1
    
    'Call FlushBuffer
    'Call WriteResuscitationToggle
End Sub

Private Sub Form_Load()
pasarme
    
    'frmMain.Caption = "Argentum Online" & " V " & App.Major & "." & _
    'App.Minor & "." & App.Revision
    
    'InvEqu.Picture = LoadPicture(App.path & _
    '"\Graficos\Centronuevoinventario.jpg")
    
   Me.Left = 0
   Me.Top = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    Select Case index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        
        Case 2
        If Command3.Visible = False Then
            pasarme
            Command3.Visible = True
        Else
            picInv.Visible = True
            hlst.Visible = False
            cmdInfo.Visible = False
            CmdLanzar.Visible = False
            Label4.Enabled = True
            Label7.Enabled = True
            Command1.Visible = False
            Command2.Visible = False
            Command3.Visible = False
            lstClases.Visible = False
        End If
        Case 3
        End
        Case 4
        'Me.Visible = False
        Me.WindowState = vbMinimized

        'Form1.Visible = True
        
        'Me.ShowInTaskbar = True
            'Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub Image3_Click(index As Integer)
    Select Case index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub



Private Sub Label4_Click()
Label4.Visible = True
Label7.Visible = True
Label4.Enabled = True
Label7.Enabled = True

    Call Audio.PlayWave(SND_CLICK)
    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevoinventario.jpg")
    picInv.Visible = True
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    lstClases.Visible = False
    Command3.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)
    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
DespInv(0).Visible = False
DespInv(1).Visible = False
End Sub

Private Sub picInv_DblClick()
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (picInv.Visible) Then
        picInv.SetFocus
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub



''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Socket1.PeerAddress)
#End If
    
    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If

    End Select
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmOldPersonaje.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False
    renderasd = False
    Call SetMusicInfo("Jugando Arduz AO - http://ao.noicoder.com/", "", "", "Games", , "{0}")
    pausa = False
    UserMeditar = False
    
#If SeguridadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If
    
        frmConnect.Show
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim Data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    'Put data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()

End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub



'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Second.Enabled = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    frmConnect.Visible = True

    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmOldPersonaje.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False
    renderasd = False
    Call SetMusicInfo("Jugando Arduz AO - http://ao.noicoder.com/", "", "", "Games", , "{0}")
    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Winsock1.RemoteHostIP)
#End If
    
    Second.Enabled = True
    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login


        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login

        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If

            

    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
#If SeguridadAlkon Then
    Call DataReceived(Data)
#End If
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If
    frmConnect.Show
End Sub
#End If

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + (32 * 17) Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + (32 * 13) Then Exit Function
    
    InGameArea = True
End Function
