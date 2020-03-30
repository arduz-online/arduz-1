VERSION 5.00
Begin VB.Form frmDebugNpc 
   Caption         =   "DebugNpcs"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   300
      Left            =   90
      TabIndex        =   5
      Top             =   2085
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ActualizarInfo"
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   1755
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "MaxNpcs:"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   1380
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "LastNpcIndex:"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1065
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Npcs Libres:"
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Npcs Activos:"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmDebugNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub Command1_Click()
Dim i As Integer, k As Integer

For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive Then k = k + 1
Next i

Label1.Caption = "Npcs Activos:" & k
Label2.Caption = "Npcs Libres:" & MAXNPCS - k
Label3.Caption = "LastNpcIndex:" & LastNPC
Label4.Caption = "MAXNPCS:" & MAXNPCS

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

