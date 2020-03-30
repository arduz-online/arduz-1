VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Audio"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Sonidos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   10
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1440
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loading As Boolean

Private Sub Check1_Click(index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    
    Select Case index
        Case 0
            If Check1(0).Value = vbUnchecked Then
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).Value = Audio.MusicVolume
            End If
        
        Case 1
            If Check1(1).Value = vbUnchecked Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
            Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).Value = Audio.SoundVolume
            End If
    End Select
End Sub





Private Sub Command2_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    

    If Audio.SoundActivated Then
        Check1(1).Value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).Value = Audio.SoundVolume
    Else
        Check1(1).Value = vbUnchecked
        Slider1(1).Enabled = False
    End If

    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub Slider1_Change(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Slider1_Scroll(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub
