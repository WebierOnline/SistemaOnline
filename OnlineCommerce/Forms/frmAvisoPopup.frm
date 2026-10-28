VERSION 5.00
Begin VB.Form frmAvisoPopup
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2200
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2200
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   0  'Manual
   Begin VB.Label lblClique
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "[CLIQUE AQUI]"
      BeginProperty Font
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1820
      Width           =   3060
   End
   Begin VB.Label lblTexto
      BackColor       =   &H00C0FFFF&
      BeginProperty Font
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   3060
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitulo
      BackColor       =   &H00C0FFFF&
      BeginProperty Font
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   2700
   End
   Begin VB.Label lblFechar
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "X"
      BeginProperty Font
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2940
      TabIndex        =   3
      Top             =   60
      Width           =   255
   End
End
Attribute VB_Name = "frmAvisoPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msTelaDestino As String

Public Sub Carregar(ByVal sTitulo As String, ByVal sTexto As String, ByVal sTelaDestino As String, ByVal lCorFundo As Long)
lblTitulo.Caption = sTitulo
lblTexto.Caption = sTexto
msTelaDestino = sTelaDestino

Me.BackColor = lCorFundo
lblTitulo.BackColor = lCorFundo
lblTexto.BackColor = lCorFundo
lblClique.BackColor = lCorFundo
lblFechar.BackColor = lCorFundo
End Sub

Private Sub lblFechar_Click()
Unload Me
End Sub

Private Sub lblClique_Click()
Select Case msTelaDestino
    Case "CONTAS_PAGAR"
        'Abre ja na aba CONSULTA, filtrado nas vencidas
        Contas_Cadastro.AbrirConsultaVencidas
        Contas_Cadastro.Show 1
    Case "CONTAS_RECEBER"
        'Abre ja filtrado em Setor=RECEBER e Situacao=VENCIDAS
        Parcelas_Consulta.AbrirConsultaVencidas
        Parcelas_Consulta.Show 1
    Case "ESTOQUE_BAIXO"
        Consulta_Estoque_Minimo.Show 1
End Select
Unload Me
End Sub
