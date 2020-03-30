VERSION 5.00
Begin VB.Form frmDebugSocket 
   Caption         =   "Debug Socket"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Reload Socket"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "State"
      Height          =   765
      Left            =   165
      TabIndex        =   8
      Top             =   4350
      Width           =   3705
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado."
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2835
      Width           =   3690
   End
   Begin VB.TextBox Text1 
      Height          =   2280
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   225
      Width           =   3690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start/Stop debug"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Errores:"
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   2610
      Width           =   2685
   End
   Begin VB.Label Label1 
      Caption         =   "Requests"
      Height          =   315
      Left            =   195
      TabIndex        =   5
      Top             =   15
      Width           =   2685
   End
End
Attribute VB_Name = "frmDebugSocket"
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
Me.Visible = False
End Sub

Private Sub Command2_Click()
DebugSocket = Not DebugSocket
End Sub

Private Sub Command3_Click()
Text1.Text = vbNullString
End Sub

Private Sub Command4_Click()
Call ReloadSokcet
End Sub
