VERSION 5.00
Begin VB.Form frmAvisoPopup
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AVISO"
   ClientHeight    =   2000
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3300
   ControlBox      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2000
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   0  'Manual
   Begin VB.Label lblExibir
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Caption         =   "[EXIBIR]"
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
      Index           =   0
      Left            =   2480
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.Label lblAlerta
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "alerta"
      BeginProperty Font
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2300
      WordWrap        =   -1  'True
   End
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
      Top             =   1650
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
      Height          =   1450
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3060
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAvisoPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msTelaDestino As String
Private msDestinos(0 To 9) As String      'destino de cada alerta ([EXIBIR]) do popup de alertas gerais

Public Sub Carregar(ByVal sTitulo As String, ByVal sTexto As String, ByVal sTelaDestino As String, ByVal lCorFundo As Long)
Me.Caption = sTitulo
lblTexto.Caption = sTexto
msTelaDestino = sTelaDestino

Me.BackColor = lCorFundo
lblTexto.BackColor = lCorFundo
lblClique.BackColor = lCorFundo
End Sub

'Popup com varias linhas, cada uma um link para um form (alertas gerais).
'sTextos = linhas separadas por vbLf; sDestinos = chaves separadas por "|" (mesma ordem)
Public Sub CarregarLinhas(ByVal sTitulo As String, ByVal sTextos As String, ByVal sDestinos As String, ByVal lCorFundo As Long)
Dim aTextos() As String
Dim aDest() As String
Dim i As Long
Dim n As Long
Dim nTop As Long

aTextos = Split(sTextos, vbLf)
aDest = Split(sDestinos, "|")
n = UBound(aTextos) + 1
If n > 10 Then n = 10

Me.Caption = sTitulo
Me.BackColor = lCorFundo
lblTexto.Visible = False
lblClique.Visible = False

'cada alerta = texto comum + [EXIBIR] no fim da linha; o popup mantem o tamanho dos outros
nTop = 120
For i = 0 To n - 1
   If i > 0 Then
      Load lblAlerta(i)
      Load lblExibir(i)
   End If
   With lblAlerta(i)
      .Caption = Replace(aTextos(i), "~", vbCrLf)          'AutoSize + WordWrap: a altura acompanha o texto
      .BackColor = lCorFundo
      .Top = nTop
      .Visible = True
   End With
   With lblExibir(i)
      .BackColor = lCorFundo
      .Top = nTop
      .Visible = True
   End With
   nTop = nTop + lblAlerta(i).Height + 150
   msDestinos(i) = aDest(i)
Next
End Sub

Private Sub lblExibir_Click(Index As Integer)
Select Case msDestinos(Index)
    Case "CONTADOR"
        'Abre ja no cadastro do contador, para preencher o e-mail
        Contador_Cadastro.AbrirPrimeiroContador
        Contador_Cadastro.Show 1
    Case "RESUMO_FISCAL"
        'abre o Resumo Fiscal ao lado do ultimo popup, na mesma altura dos popups
        Notas_Adesivas.Show
        PosicionarAoLadoDosPopups Notas_Adesivas
    Case "EMPRESA"
        Empresa_Cadastro.Show
    Case "ENTRADA_ESTOQUE"
        Entrada_Estoque.Show
    Case "NOTAS_DESTINADAS"
        Notas_Destinadas_Consulta.Show 1
End Select
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
End Sub
