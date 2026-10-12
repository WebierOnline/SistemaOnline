VERSION 5.00
Begin VB.Form PDV_PixVenda
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento com Pix"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPix
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6600
      Top             =   120
   End
   Begin VB.Timer tmrFechar
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   6120
      Top             =   120
   End
   Begin VB.CommandButton cmdNovo
      Caption         =   "Gerar novo QR Code"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   6900
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.CommandButton cmdCancelar
      Cancel          =   -1  'True
      Caption         =   "Cancelar (Esc)"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      TabIndex        =   1
      Top             =   6900
      Width           =   3240
   End
   Begin VB.Image imgQR
      Height          =   4200
      Left            =   1500
      Stretch         =   -1  'True
      Top             =   1500
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Label lblTitulo
      Alignment       =   2  'Center
      Caption         =   "PAGAMENTO COM PIX"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   6720
   End
   Begin VB.Label lblValor
      Alignment       =   2  'Center
      Caption         =   "R$ 0,00"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   690
      Left            =   240
      TabIndex        =   3
      Top             =   660
      Width           =   6720
   End
   Begin VB.Label lblInstrucao
      Alignment       =   2  'Center
      Caption         =   "Peça ao cliente para ler o QR Code no aplicativo do banco."
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   6720
   End
   Begin VB.Label lblStatus
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   6720
   End
End
Attribute VB_Name = "PDV_PixVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' Pix dinâmico na venda: mostra o QR Code da cobrança (criada pelo servidor na conta Efí
' do restaurante) e fecha sozinha quando o banco confirma o pagamento.
' Quem chama: PDV.ReceberPixDinamico (lê pPago e descarrega).
'=============================================================================
Option Explicit

Public pPago As Boolean
Public pTxId As String

Private mValor As Currency
Private mPedido As String
Private mCaixa As String
Private mExpira As Date
Private mOcupado As Boolean
Private mIniciado As Boolean

Public Sub Preparar(ByVal Valor As Currency, ByVal Pedido As String, ByVal Caixa As String)
   mValor = Valor
   mPedido = Pedido
   mCaixa = Caixa
   pPago = False
   pTxId = ""
   lblValor.Caption = "R$ " & Format$(Valor, "#,##0.00")
End Sub

Private Sub Form_Activate()
   If mIniciado Then Exit Sub
   mIniciado = True
   Gerar
End Sub

Private Sub Gerar()
   Dim copia As String, arquivo As String, pago As Boolean, msg As String, ok As Boolean

   If mOcupado Then Exit Sub
   mOcupado = True
   tmrPix.Enabled = False
   cmdNovo.Visible = False
   imgQR.Visible = False
   lblStatus.ForeColor = &HC0&
   lblStatus.Caption = "Gerando o QR Code... AGUARDE!"
   Screen.MousePointer = vbHourglass
   DoEvents
   ok = LicPixVendaCriar(mValor, mPedido, mCaixa, pTxId, copia, arquivo, pago, msg)
   Screen.MousePointer = vbDefault
   mOcupado = False
   If Not ok Then
      lblStatus.Caption = msg
      cmdNovo.Caption = "Tentar de novo"
      cmdNovo.Visible = True
      Exit Sub
   End If
   If pago Then
      Confirmar
      Exit Sub
   End If
   If Len(arquivo) > 0 Then
      On Error Resume Next
      Set imgQR.Picture = LoadPicture(arquivo)
      imgQR.Visible = (Err.Number = 0)
      On Error GoTo 0
   End If
   'O servidor dá 15 minutos; aqui expira 1 minuto antes.
   mExpira = DateAdd("n", 14, Now)
   lblStatus.ForeColor = &H800000
   lblStatus.Caption = "Aguardando o pagamento..."
   tmrPix.Enabled = True
End Sub

Private Sub tmrPix_Timer()
   Dim s As Long

   If mOcupado Or Len(pTxId) = 0 Then Exit Sub
   mOcupado = True
   s = LicPixVendaSituacao(pTxId)
   mOcupado = False
   If s = 1 Then
      Confirmar
   ElseIf Now > mExpira Then
      tmrPix.Enabled = False
      imgQR.Visible = False
      lblStatus.ForeColor = &HC0&
      lblStatus.Caption = "O QR Code expirou. Gere outro."
      cmdNovo.Caption = "Gerar novo QR Code"
      cmdNovo.Visible = True
   ElseIf s = -1 Then
      lblStatus.Caption = "Aguardando o pagamento... (sem resposta do servidor, tentando de novo)"
   Else
      lblStatus.Caption = "Aguardando o pagamento...  (expira em " & (DateDiff("s", Now, mExpira) \ 60) + 1 & " min)"
   End If
End Sub

Private Sub Confirmar()
   tmrPix.Enabled = False
   pPago = True
   imgQR.Visible = False
   cmdNovo.Visible = False
   cmdCancelar.Enabled = False
   lblStatus.ForeColor = &H8000&
   lblStatus.Caption = "PAGAMENTO CONFIRMADO!"
   Beep
   tmrFechar.Enabled = True
End Sub

Private Sub tmrFechar_Timer()
   tmrFechar.Enabled = False
   Me.Hide
End Sub

Private Sub cmdNovo_Click()
   Gerar
End Sub

Private Sub cmdCancelar_Click()
   Dim s As Long

   If mOcupado Or pPago Then Exit Sub
   If Len(pTxId) > 0 Then
      'Timer parado enquanto pergunta (sem fechar a tela por trás da mensagem).
      tmrPix.Enabled = False
      If MsgBox("Cancelar o Pix?" & vbCrLf & vbCrLf & _
                "O QR Code deixa de valer e o cliente não consegue mais pagá-lo.", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Pix") = vbNo Then
         tmrPix.Enabled = Not cmdNovo.Visible
         Exit Sub
      End If
      'Invalida o QR no banco. Se o cliente pagou neste instante, o banco recusa e a venda é finalizada.
      mOcupado = True
      Screen.MousePointer = vbHourglass
      s = LicPixVendaCancelar(pTxId)
      Screen.MousePointer = vbDefault
      mOcupado = False
      If s = 1 Then
         Confirmar
         Exit Sub
      End If
   End If
   tmrPix.Enabled = False
   pPago = False
   Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Só sai pelos botões.
   If UnloadMode = vbFormControlMenu Then Cancel = True
End Sub
