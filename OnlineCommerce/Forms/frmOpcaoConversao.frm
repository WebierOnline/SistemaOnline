VERSION 5.00
Begin VB.Form frmOpcaoConversao
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Converter Pedidos"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNFSeparada
      Caption         =   "NF SEPARADA"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdNFAgrupada
      Caption         =   "NF AGRUPADA"
      Height          =   615
      Left            =   2100
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar
      Cancel          =   -1  'True
      Caption         =   "CANCELAR"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblDesc3
      Caption         =   "Não converter"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblDesc2
      Caption         =   "Todos em uma NF"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2100
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblDesc1
      Caption         =   "Uma NF por pedido"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPergunta
      Caption         =   ""
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmOpcaoConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Resultado As Integer

Private Sub cmdNFSeparada_Click()
    Resultado = 1
    Unload Me
End Sub

Private Sub cmdNFAgrupada_Click()
    Resultado = 2
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Resultado = 0
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Resultado = 0
    End If
End Sub
