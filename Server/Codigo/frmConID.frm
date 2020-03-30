VERSION 5.00
Begin VB.Form frmConID 
   Caption         =   "ConID"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Liberar todos los slots"
      Height          =   390
      Left            =   135
      TabIndex        =   3
      Top             =   3495
      Width           =   4290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver estado"
      Height          =   390
      Left            =   135
      TabIndex        =   2
      Top             =   3030
      Width           =   4290
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   3975
      Width           =   4290
   End
   Begin VB.Label Label1 
      Height          =   510
      Left            =   180
      TabIndex        =   4
      Top             =   2430
      Width           =   4230
   End
End
Attribute VB_Name = "frmConID"
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
Unload Me
End Sub

Private Sub Command2_Click()

List1.Clear

Dim c As Integer
Dim i As Integer

For i = 1 To MaxUsers
    List1.AddItem "UserIndex " & i & " -- " & UserList(i).ConnID
    If UserList(i).ConnID <> -1 Then c = c + 1
Next i

If c = MaxUsers Then
    Label1.Caption = "No hay slots vacios!"
Else
    Label1.Caption = "Hay " & MaxUsers - c & " slots vacios!"
End If

End Sub

Private Sub Command3_Click()
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
Next i

End Sub

