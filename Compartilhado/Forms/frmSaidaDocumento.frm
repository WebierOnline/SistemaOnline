VERSION 5.00
Begin VB.Form frmSaidaDocumento
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saída do documento"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame1
      Caption         =   " Escolha como sair "
      Height          =   2055
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optWhatsapp
         Caption         =   "Enviar por WhatsApp"
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   3015
      End
      Begin VB.OptionButton optEmail
         Caption         =   "Enviar por email"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   1140
         Width           =   3015
      End
      Begin VB.OptionButton optPDF
         Caption         =   "Gerar PDF"
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton optImprimir
         Caption         =   "Imprimir"
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmSaidaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Tela modal de escolha da saida do documento.
'Escolha: 0 = cancelou, 1 = imprimir, 2 = gerar PDF, 3 = enviar por e-mail, 4 = WhatsApp
Public Escolha As Integer

Private Sub Form_Load()
Escolha = 0
optImprimir.Value = True
End Sub

Private Sub cmdOK_Click()
If optImprimir.Value Then
    Escolha = 1
ElseIf optPDF.Value Then
    Escolha = 2
ElseIf optEmail.Value Then
    Escolha = 3
ElseIf optWhatsapp.Value Then
    Escolha = 4
End If
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
Escolha = 0
Me.Hide
End Sub

'Pergunta a saida do documento. Retorna 0 se cancelou.
'bSoEnvio = so as opcoes de envio (e-mail e WhatsApp), ex.: Enviar XML da NF-e.
Public Function Perguntar(Optional ByVal bSoEnvio As Boolean = False) As Integer
Escolha = 0
optImprimir.Enabled = Not bSoEnvio
optPDF.Enabled = Not bSoEnvio
If bSoEnvio Then optEmail.Value = True Else optImprimir.Value = True
Me.Show vbModal
Perguntar = Escolha
End Function
