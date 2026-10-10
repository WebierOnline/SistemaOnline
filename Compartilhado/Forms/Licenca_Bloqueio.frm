VERSION 5.00
Begin VB.Form Licenca_Bloqueio
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licença do sistema"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPix
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9000
      Top             =   120
   End
   Begin VB.TextBox txtCodigo
      BeginProperty Font
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      MaxLength       =   9
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdCodigo
      Caption         =   "LIBERAR"
      Height          =   420
      Left            =   4320
      TabIndex        =   1
      Top             =   5512
      Width           =   1455
   End
   Begin VB.CommandButton cmdPagar
      Caption         =   "PAGAR AGORA (PIX)"
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
      TabIndex        =   2
      Top             =   6780
      Width           =   2895
   End
   Begin VB.CommandButton cmdVerificar
      Caption         =   "VERIFICAR NOVAMENTE"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3352
      TabIndex        =   3
      Top             =   6780
      Width           =   2895
   End
   Begin VB.CommandButton cmdFechar
      Caption         =   "FECHAR O SISTEMA"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6465
      TabIndex        =   4
      Top             =   6780
      Width           =   2895
   End
   Begin VB.CommandButton cmdCopiar
      Caption         =   "Copiar código Pix"
      Height          =   420
      Left            =   3240
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtCopiaCola
      Height          =   900
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3780
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Image imgQR
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblTitulo
      Caption         =   "SISTEMA BLOQUEADO"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   9120
   End
   Begin VB.Label lblMensagem
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   9120
   End
   Begin VB.Label lblReferencia
      BeginProperty Font
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   3240
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label lblPix
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3240
      TabIndex        =   9
      Top             =   2940
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label lblCodigo
      Caption         =   "Código de liberação:"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   10
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   5580
      Width           =   2055
   End
   Begin VB.Label lblStatus
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
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   6120
      Width           =   9120
   End
End
Attribute VB_Name = "Licenca_Bloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
' Tela de aviso de vencimento / bloqueio da licença online (OnlineCommerce, PDV, OS)
' e de pagamento antecipado da mensalidade (botão "Licença" do PDV).
' Pagar agora (Pix) libera sozinho quando o pagamento é confirmado.
' Quem chama: LicencaOnline.VerificarLicencaOnline / LicPagarMensalidade (lê pDesbloqueado e descarrega).
'=============================================================================
Option Explicit

Public pDesbloqueado As Boolean
Public pModoAviso As Boolean

Private mModoPagar As Boolean
Private mCodUsuario As Long
Private mTxId As String

'ModoAviso = só avisa (botão CONTINUAR); senão bloqueia (botão FECHAR O SISTEMA).
Public Sub Preparar(ByVal ModoAviso As Boolean, ByVal CodUsuario As Long)
   pModoAviso = ModoAviso
   mModoPagar = False
   pDesbloqueado = False
   mCodUsuario = CodUsuario
   mTxId = ""
   AtualizarTela
End Sub

'Só pagar a mensalidade (antes do vencimento); fechar não afeta o sistema.
Public Sub PrepararPagamento(ByVal CodUsuario As Long)
   pModoAviso = True
   mModoPagar = True
   pDesbloqueado = True
   mCodUsuario = CodUsuario
   mTxId = ""
   AtualizarTela
End Sub

Private Sub AtualizarTela()
   If mModoPagar Then
      Me.Caption = "Licença do sistema - pagar mensalidade"
      lblTitulo.Caption = "PAGAR MENSALIDADE"
      lblTitulo.ForeColor = &H800000
      cmdFechar.Caption = "FECHAR"
      lblMensagem.Caption = LicMensagemPagamento()
      cmdPagar.Enabled = True
      lblCodigo.Visible = False
      txtCodigo.Visible = False
      cmdCodigo.Visible = False
      Exit Sub
   End If
   If pModoAviso Then
      Me.Caption = "Licença do sistema - aviso"
      lblTitulo.Caption = "AVISO DE VENCIMENTO"
      lblTitulo.ForeColor = &H4080&
      cmdFechar.Caption = "CONTINUAR"
   Else
      Me.Caption = "Licença do sistema - bloqueado"
      lblTitulo.Caption = "SISTEMA BLOQUEADO"
      lblTitulo.ForeColor = &HC0&
      cmdFechar.Caption = "FECHAR O SISTEMA"
   End If
   lblMensagem.Caption = LicMensagemTela()
   cmdPagar.Enabled = (LicEstadoAtual = licBloqueada Or LicEstadoAtual = licAviso)
   'Código de liberação só faz sentido bloqueado.
   lblCodigo.Visible = Not pModoAviso
   txtCodigo.Visible = Not pModoAviso
   cmdCodigo.Visible = Not pModoAviso
End Sub

Private Sub cmdPagar_Click()
   Dim copiaECola As String, arquivo As String, descricao As String, msg As String

   Screen.MousePointer = vbHourglass
   lblStatus.Caption = "Gerando o Pix..."
   DoEvents
   If LicCriarPix(mTxId, copiaECola, arquivo, descricao, msg) Then
      If Len(arquivo) > 0 Then
         On Error Resume Next
         Set imgQR.Picture = LoadPicture(arquivo)
         imgQR.Visible = (Err.Number = 0)
         On Error GoTo 0
      End If
      lblReferencia.Caption = descricao
      lblReferencia.Visible = True
      lblPix.Caption = "Pague pelo aplicativo do banco (QR Code ou Pix copia e cola)." & vbCrLf & _
                       IIf(mModoPagar, "Assim que o pagamento for confirmado, a licença é atualizada sozinha.", _
                                       "Assim que o pagamento for confirmado, o sistema libera sozinho.")
      lblPix.Visible = True
      txtCopiaCola.Text = copiaECola
      txtCopiaCola.Visible = True
      cmdCopiar.Visible = True
      lblStatus.Caption = "Aguardando a confirmação do pagamento..."
      tmrPix.Enabled = True
   Else
      lblStatus.Caption = msg
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCopiar_Click()
   Clipboard.Clear
   Clipboard.SetText txtCopiaCola.Text
   lblStatus.Caption = "Código Pix copiado. Cole no aplicativo do banco (Pix copia e cola)."
End Sub

Private Sub tmrPix_Timer()
   If Len(mTxId) = 0 Then tmrPix.Enabled = False: Exit Sub
   tmrPix.Enabled = False
   If LicVerificarPix(mTxId) Then
      If PodeLiberar(LicAvaliar()) Or mModoPagar Then
         Liberar "Pagamento confirmado. Obrigado!"
         Exit Sub
      End If
      AtualizarTela
   End If
   tmrPix.Enabled = True
End Sub

Private Sub cmdVerificar_Click()
   Dim msg As String

   Screen.MousePointer = vbHourglass
   lblStatus.Caption = "Consultando o servidor de licenças..."
   DoEvents
   'Pix gerado nesta tela: confere primeiro o pagamento.
   If Len(mTxId) > 0 Then
      If LicVerificarPix(mTxId) And mModoPagar Then
         Screen.MousePointer = vbDefault
         Liberar "Pagamento confirmado. Obrigado!"
         Exit Sub
      End If
   End If
   LicSincronizar msg
   Screen.MousePointer = vbDefault

   If mModoPagar Then
      LicAvaliar
      AtualizarTela
      lblStatus.Caption = IIf(Len(msg) > 0, msg, IIf(Len(mTxId) > 0, "Pagamento ainda não confirmado.", "Licença atualizada."))
      Exit Sub
   End If

   If PodeLiberar(LicAvaliar()) Then
      Liberar "Licença em dia. O sistema foi liberado."
   Else
      AtualizarTela
      lblStatus.Caption = IIf(Len(msg) > 0, msg, "A licença continua pendente.")
   End If
End Sub

Private Sub cmdCodigo_Click()
   Dim msg As String

   If Len(Trim$(txtCodigo.Text)) = 0 Then
      lblStatus.Caption = "Digite o código de liberação."
      txtCodigo.SetFocus
      Exit Sub
   End If
   If LicAplicarCodigo(txtCodigo.Text, mCodUsuario, msg) Then
      If PodeLiberar(LicAvaliar()) Then
         Liberar msg
         Exit Sub
      End If
      AtualizarTela
   End If
   lblStatus.Caption = msg
   txtCodigo.SelStart = 0
   txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      cmdCodigo_Click
   Else
      KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
   End If
End Sub

Private Sub cmdFechar_Click()
   tmrPix.Enabled = False
   'No aviso, continuar usando; no bloqueio, quem chamou fecha o sistema.
   pDesbloqueado = pModoAviso
   Me.Hide
End Sub

Private Function PodeLiberar(ByVal Estado As LicEstado) As Boolean
   PodeLiberar = (Estado = licLiberada Or Estado = licAviso)
End Function

Private Sub Liberar(ByVal Mensagem As String)
   tmrPix.Enabled = False
   pDesbloqueado = True
   MsgBox Mensagem, vbInformation, "Licença do sistema"
   Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Só sai pelos botões (Alt+F4 não pula o bloqueio).
   If UnloadMode = vbFormControlMenu Then Cancel = True
End Sub
