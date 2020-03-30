VERSION 5.00
Begin VB.Form frmTrafic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trafico"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   2250
      Width           =   960
   End
   Begin VB.ListBox lstTrafico 
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   4455
   End
End
Attribute VB_Name = "frmTrafic"
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
