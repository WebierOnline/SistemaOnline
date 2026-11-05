VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Begin VB.Form Licenca_Bloqueio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licença do sistema"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPix 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9000
      Top             =   120
   End
   Begin VB.Timer tmrEspera 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8520
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
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox txtCopiaCola 
      Height          =   900
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4860
      Visible         =   0   'False
      Width           =   6120
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   615
      Left            =   6180
      TabIndex        =   8
      Top             =   7860
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "FECHAR O SISTEMA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Licenca_Bloqueio.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdPagar 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   7860
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "PAGAR AGORA (PIX)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   128
      FCOLO           =   128
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Licenca_Bloqueio.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdVerificar 
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   7860
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "VERIFICAR NOVAMENTE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Licenca_Bloqueio.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCopiar 
      Height          =   435
      Left            =   3240
      TabIndex        =   11
      Top             =   5820
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Copiar código Pix"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Licenca_Bloqueio.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdCodigo 
      Height          =   435
      Left            =   4260
      TabIndex        =   12
      Top             =   6540
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "LIBERAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Licenca_Bloqueio.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgQR 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3600
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
      TabIndex        =   2
      Top             =   180
      Width           =   9120
   End
   Begin VB.PictureBox picMensagem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   9120
      TabIndex        =   3
      TabStop         =   0   'False
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
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label lblPix 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3240
      TabIndex        =   4
      Top             =   4020
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Código de liberação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   6660
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
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   7170
      Width           =   9120
   End
   Begin VB.Shape shpTrilho 
      BorderColor     =   &H00C0C0C0&
      Height          =   210
      Left            =   240
      Top             =   7530
      Visible         =   0   'False
      Width           =   9120
   End
   Begin VB.Shape shpBarra 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   270
      Top             =   7560
      Visible         =   0   'False
      Width           =   1800
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
Private mOcupado As Boolean
Private mTextoEspera As String
Private mPagarEstava As Boolean
Private mCodUsuario As Long
Private mTxId As String
Private mGerarAoAbrir As Boolean

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
'GerarPix = gera o QR Code assim que a tela aparece (sem precisar clicar em PAGAR AGORA).
Public Sub PrepararPagamento(ByVal CodUsuario As Long, Optional ByVal GerarPix As Boolean = False)
   pModoAviso = True
   mModoPagar = True
   pDesbloqueado = True
   mCodUsuario = CodUsuario
   mTxId = ""
   mGerarAoAbrir = GerarPix
   AtualizarTela
End Sub

Private Sub AtualizarTela()
   If mModoPagar Then
      Me.Caption = "Licença do sistema - pagar mensalidade"
      lblTitulo.Caption = "PAGAR MENSALIDADE"
      lblTitulo.ForeColor = &H800000
      cmdFechar.Caption = "FECHAR"
      DesenharMensagem LicMensagemPagamento()
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
   DesenharMensagem LicMensagemTela()
   cmdPagar.Enabled = (LicEstadoAtual = licBloqueada Or LicEstadoAtual = licAviso)
   'Código de liberação só faz sentido bloqueado.
   lblCodigo.Visible = Not pModoAviso
   txtCodigo.Visible = Not pModoAviso
   cmdCodigo.Visible = Not pModoAviso
End Sub

'Mensagem com negrito no meio da linha (o Label não faz): desenhada no picMensagem.
'**texto** = negrito; linha vazia = espaço entre parágrafos; linha "• ..." quebra alinhada depois do marcador.
Private Sub DesenharMensagem(ByVal Texto As String)
   Dim linhas() As String, partes() As String, palavras() As String, palavra As String
   Dim i As Long, j As Long, k As Long, alturaLinha As Single, recuo As Single

   picMensagem.BackColor = Me.BackColor
   picMensagem.Cls
   picMensagem.FontBold = False
   alturaLinha = picMensagem.TextHeight("Ag")
   picMensagem.CurrentY = 0
   linhas = Split(Replace(Texto, vbCr, ""), vbLf)
   For i = 0 To UBound(linhas)
      If Len(Trim$(linhas(i))) = 0 Then
         picMensagem.CurrentY = picMensagem.CurrentY + alturaLinha * 0.6
      Else
         recuo = 0
         If Left$(linhas(i), 2) = "• " Then recuo = picMensagem.TextWidth("• ")
         picMensagem.CurrentX = 0
         partes = Split(linhas(i), "**")
         For j = 0 To UBound(partes)
            picMensagem.FontBold = (j Mod 2 = 1)
            palavras = Split(partes(j), " ")
            For k = 0 To UBound(palavras)
               palavra = palavras(k) & IIf(k < UBound(palavras), " ", "")
               'Não cabe no resto da linha: continua na de baixo (sem começar por espaço).
               If Len(palavra) > 0 And picMensagem.CurrentX > recuo Then
                  If picMensagem.CurrentX + picMensagem.TextWidth(RTrim$(palavra)) > picMensagem.ScaleWidth Then
                     picMensagem.CurrentY = picMensagem.CurrentY + alturaLinha
                     picMensagem.CurrentX = recuo
                     If palavra = " " Then palavra = ""
                  End If
               End If
               If Len(palavra) > 0 Then picMensagem.Print palavra;
            Next
         Next
         picMensagem.FontBold = False
         picMensagem.CurrentY = picMensagem.CurrentY + alturaLinha
      End If
   Next
End Sub

'Enquanto fala com o servidor: botões desligados e "AGUARDE!" com os segundos (a tela não parece travada).
Private Sub Ocupar(ByVal texto As String)
   mOcupado = True
   mTextoEspera = texto
   mPagarEstava = cmdPagar.Enabled
   cmdPagar.Enabled = False
   cmdVerificar.Enabled = False
   cmdFechar.Enabled = False
   cmdCodigo.Enabled = False
   lblStatus.ForeColor = &HC0&
   lblStatus.Caption = texto & " AGUARDE!"
   Screen.MousePointer = vbHourglass
   shpBarra.Left = shpTrilho.Left + 30
   shpTrilho.Visible = True
   shpBarra.Visible = True
   tmrEspera.Enabled = True
   DoEvents
End Sub

Private Sub Desocupar()
   tmrEspera.Enabled = False
   shpTrilho.Visible = False
   shpBarra.Visible = False
   lblStatus.ForeColor = &H800000
   cmdPagar.Enabled = mPagarEstava
   cmdVerificar.Enabled = True
   cmdFechar.Enabled = True
   cmdCodigo.Enabled = True
   Screen.MousePointer = vbDefault
   mOcupado = False
End Sub

'Barra andando da esquerda para a direita enquanto aguarda (mostra que não travou).
Private Sub tmrEspera_Timer()
   Dim texto As String

   If shpBarra.Left + 150 > shpTrilho.Left + shpTrilho.Width - shpBarra.Width - 30 Then
      shpBarra.Left = shpTrilho.Left + 30
   Else
      shpBarra.Left = shpBarra.Left + 150
   End If
   texto = mTextoEspera & " AGUARDE!" & IIf(LicSegundosEspera > 0, "  (" & LicSegundosEspera & "s)", "")
   If lblStatus.Caption <> texto Then lblStatus.Caption = texto
End Sub

Private Sub cmdPagar_Click()
   Dim copiaECola As String, arquivo As String, descricao As String, msg As String, ok As Boolean

   If mOcupado Then Exit Sub
   Ocupar "Gerando o Pix (QR Code)..."
   ok = LicCriarPix(mTxId, copiaECola, arquivo, descricao, msg)
   Desocupar
   If ok Then
      If Len(arquivo) > 0 Then
         On Error Resume Next
         Set imgQR.Picture = LoadPicture(arquivo)
         imgQR.Visible = (Err.Number = 0)
         On Error GoTo 0
      End If
      lblReferencia.Caption = descricao
      AjustarReferencia
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
End Sub

'Referência numa linha só: diminui a fonte (12 até 8,25) até caber. Várias mensalidades ou multa e juros deixam o texto longo.
Private Sub AjustarReferencia()
   Dim tamanho As Single

   Me.FontName = lblReferencia.FontName
   Me.FontBold = lblReferencia.FontBold
   For tamanho = 12 To 8.25 Step -0.75
      Me.FontSize = tamanho
      If Me.TextWidth(lblReferencia.Caption) <= lblReferencia.Width - 60 Then Exit For
   Next
   lblReferencia.FontSize = Me.FontSize
End Sub

Private Sub cmdCopiar_Click()
   Clipboard.Clear
   Clipboard.SetText txtCopiaCola.Text
   lblStatus.Caption = "Código Pix copiado. Cole no aplicativo do banco (Pix copia e cola)."
End Sub

Private Sub tmrPix_Timer()
   Dim pago As Boolean

   If Len(mTxId) = 0 Then tmrPix.Enabled = False: Exit Sub
   If mOcupado Then Exit Sub
   tmrPix.Enabled = False
   'Consulta em segundo plano: não mexe na tela, só impede clique duplo enquanto consulta.
   mOcupado = True
   pago = LicVerificarPix(mTxId)
   mOcupado = False
   If pago Then
      If PodeLiberar(LicAvaliar()) Or mModoPagar Then
         Liberar "Pagamento confirmado. Obrigado!"
         Exit Sub
      End If
      AtualizarTela
   End If
   tmrPix.Enabled = True
End Sub

Private Sub cmdVerificar_Click()
   Dim msg As String, pago As Boolean

   If mOcupado Then Exit Sub
   Ocupar "Consultando o servidor de licenças..."
   'Pix gerado nesta tela: confere primeiro o pagamento.
   If Len(mTxId) > 0 Then pago = LicVerificarPix(mTxId)
   If pago And mModoPagar Then
      Desocupar
      Liberar "Pagamento confirmado. Obrigado!"
      Exit Sub
   End If
   LicSincronizar msg
   Desocupar

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

   If mOcupado Then Exit Sub
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
   If mOcupado Then Exit Sub
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

Private Sub Form_Activate()
   If Not mGerarAoAbrir Then Exit Sub
   mGerarAoAbrir = False
   DoEvents
   cmdPagar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Só sai pelos botões (Alt+F4 não pula o bloqueio).
   If UnloadMode = vbFormControlMenu Then Cancel = True
End Sub
