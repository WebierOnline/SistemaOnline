VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCobrancaEnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobrança por WhatsApp"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   16815
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2115
      Left            =   3420
      TabIndex        =   20
      Top             =   1980
      Width           =   2955
      ExtentX         =   5212
      ExtentY         =   3731
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Histórico de envios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   120
      TabIndex        =   13
      Top             =   5340
      Width           =   16635
      Begin VB.ComboBox cboPeriodo 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid grdHist 
         Height          =   3225
         Left            =   60
         TabIndex        =   15
         Top             =   600
         Width           =   16500
         _ExtentX        =   29104
         _ExtentY        =   5689
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar por:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   16725
      TabIndex        =   10
      Top             =   0
      Width           =   16755
      Begin VB.Label lblResumo 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   7380
         TabIndex        =   12
         Top             =   240
         Width           =   9240
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrança por WhatsApp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1620
         TabIndex        =   11
         Top             =   240
         Width           =   3750
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   540
         Picture         =   "frmCobrancaEnvio.frx":0000
         Top             =   60
         Width           =   675
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdEnvio 
      Height          =   3540
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   6244
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.TextBox txtPrevia 
      Height          =   3645
      Left            =   12180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   4560
   End
   Begin ChamaleonBtn.chameleonButton cmdMarcarTodos 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   4860
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Marcar todos"
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
      MICON           =   "frmCobrancaEnvio.frx":0C70
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdDesmarcar 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4860
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Desmarcar todos"
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
      MICON           =   "frmCobrancaEnvio.frx":0C8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdNaoCobrar 
      Height          =   375
      Left            =   2940
      TabIndex        =   4
      Top             =   4860
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Não cobrar este cliente"
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
      MICON           =   "frmCobrancaEnvio.frx":0CA8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdEnviar 
      Height          =   375
      Left            =   13200
      TabIndex        =   5
      Top             =   4800
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Enviar marcados"
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
      MICON           =   "frmCobrancaEnvio.frx":0CC4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdParar 
      Height          =   375
      Left            =   14760
      TabIndex        =   6
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Parar"
      ENAB            =   0   'False
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
      MICON           =   "frmCobrancaEnvio.frx":0CE0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizar 
      Height          =   375
      Left            =   15780
      TabIndex        =   7
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Atualizar"
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
      MICON           =   "frmCobrancaEnvio.frx":0CFC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdRegua 
      Height          =   375
      Left            =   13500
      TabIndex        =   17
      Top             =   9360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Régua de cobrança..."
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
      MICON           =   "frmCobrancaEnvio.frx":0D18
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdFechar 
      Height          =   375
      Left            =   15540
      TabIndex        =   18
      Top             =   9360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Fechar"
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
      MICON           =   "frmCobrancaEnvio.frx":0D34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonBtn.chameleonButton cmdAtualizarCliente 
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   4860
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Atualizar Cliente"
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
      MICON           =   "frmCobrancaEnvio.frx":0D50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkJuros 
      Caption         =   "Cobrar juros e multa"
      Height          =   255
      Left            =   7260
      TabIndex        =   21
      ToolTipText     =   "Marcado: o valor cobrado inclui multa e juros do atraso (Configuração Geral). A mensagem não fala em juros, só mostra o valor."
      Top             =   4860
      Value           =   1  'Checked
      Width           =   5895
   End
   Begin VB.OptionButton optFiltro 
      Caption         =   "À vencer"
      Height          =   255
      Index           =   1
      Left            =   900
      TabIndex        =   23
      Top             =   900
      Width           =   1935
   End
   Begin VB.OptionButton optFiltro 
      Caption         =   "Vencendo hoje"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   24
      Top             =   900
      Width           =   1575
   End
   Begin VB.OptionButton optFiltro 
      Caption         =   "Vencido"
      Height          =   255
      Index           =   3
      Left            =   4500
      TabIndex        =   25
      Top             =   900
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Image ImgMarcada 
      Height          =   195
      Left            =   7320
      Picture         =   "frmCobrancaEnvio.frx":0D6C
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDesmarcada 
      Height          =   195
      Left            =   7620
      Picture         =   "frmCobrancaEnvio.frx":316B
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblMostrar 
      Caption         =   "Mostrar:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   930
      Width           =   735
   End
   Begin VB.Label lblPrevia 
      Caption         =   "Mensagem (prévia):"
      Height          =   255
      Left            =   12180
      TabIndex        =   8
      Top             =   840
      Width           =   3540
   End
   Begin VB.Label lblProgresso 
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   9450
      Width           =   13200
   End
End
Attribute VB_Name = "frmCobrancaEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Cobrança por WhatsApp: quem recebe hoje pela régua (modCobranca.Cob_Pendentes), envio manual dos marcados
'Mostrar: À vencer / Vencendo hoje / Vencido (padrão) = todos os clientes da situação; quem não tem envio hoje
'aparece em cinza com a data prevista, para dar tempo de marcar "Não cobrar" antes do envio (inclusive o automático).
'(com o intervalo e o limite da configuração) e histórico. A régua e a configuração ficam em frmCobrancaRegua.
'O GIF "aguarde" vem embutido no executável (OnlineCommerce.res, recurso "GIF" 101) e é mostrado no WebBrowser1.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const TITULO As String = "Cobrança por WhatsApp"

Private Const C_MARCA As Integer = 0
Private Const C_CLIENTE As Integer = 1
Private Const C_CELULAR As Integer = 2
Private Const C_REGRA As Integer = 3
Private Const C_PARC As Integer = 4
Private Const C_TOTAL As Integer = 5
Private Const C_ATRASO As Integer = 6
Private Const C_SITUACAO As Integer = 7

Private mItens() As CobItem
Private mMarca() As Boolean       'por item (índice de mItens)
Private mQtd As Long
Private mCfg As CobConfig
Private mEnviando As Boolean
Private mIniciou As Boolean
Private mSemEvento As Boolean     'mudando Row/Col por código: não atualiza a prévia
Private mAguardeHtml As String
Private mAguardeCarregado As Boolean
Private mCarregandoLista As Boolean   'Carregar em andamento (o GIF anima com DoEvents: bloqueia cliques)
Private mForaDoPlano As Boolean

Private Sub Form_Load()
mIniciou = False
mEnviando = False
WebBrowser1.Visible = False
cboPeriodo.AddItem "Hoje"
cboPeriodo.AddItem "Últimos 7 dias"
cboPeriodo.AddItem "Últimos 30 dias"
cboPeriodo.ListIndex = 0
lblProgresso.Caption = ""
lblResumo.Caption = ""
txtPrevia.Text = ""
LegendaJuros
optFiltro(COB_FILTRO_A_VENCER).Caption = "À vencer (" & COB_DIAS_A_VENCER & " dias)"
End Sub

'Situação mostrada no grid (padrão: Vencido)
Private Function FiltroAtual() As Integer
Dim k As Integer

FiltroAtual = COB_FILTRO_VENCIDO
For k = COB_FILTRO_A_VENCER To COB_FILTRO_VENCIDO
    If optFiltro(k).Value Then FiltroAtual = k
Next
End Function

Private Sub optFiltro_Click(Index As Integer)
If mIniciou And Not mEnviando And Not mCarregandoLista Then Carregar
End Sub

'Mostra a multa e o juros ao dia da Configuração Geral no próprio checkbox
Private Sub LegendaJuros()
Dim cJurosDia As Currency, nTipo As Integer, cMulta As Currency

Cob_JurosConfig cJurosDia, nTipo, cMulta
If cJurosDia > 0 Or cMulta > 0 Then
    chkJuros.Caption = "Cobrar juros e multa (multa " & Format$(cMulta, "0.00") & "% + " & Format$(cJurosDia, "0.00##") & "% ao dia, " & _
                       IIf(nTipo = 1, "sobre a parcela", "sobre o saldo") & ")"
Else
    chkJuros.Caption = "Cobrar juros e multa (sem multa e juros na Configuração Geral)"
End If
End Sub

Private Sub chkJuros_Click()
If mIniciou And Not mEnviando And Not mCarregandoLista Then Recarregar
End Sub

Private Sub Form_Activate()
If mIniciou Then Exit Sub
mIniciou = True
If Not Cob_BancoPronto() Then
    MsgBox "O banco de dados ainda não tem as tabelas da cobrança por WhatsApp (script 184)." & vbCrLf & vbCrLf & _
           "Atualize o sistema e tente de novo.", vbExclamation, TITULO
    Unload Me
    Exit Sub
End If
Cob_GarantirRegrasPadrao
Carregar
CarregarHistorico
VerificarWhatsApp
End Sub

'Avisa logo ao abrir se o WhatsApp não está no plano ou não está conectado (quem decide é o servidor)
Private Sub VerificarWhatsApp()
Dim sSit As String, sNum As String, sNome As String, sMsg As String

Me.MousePointer = vbHourglass
sSit = ZapOnlineSituacao(sNum, sNome, sMsg)
Me.MousePointer = vbDefault
If ZapOnlineForaDoPlano() Then
    MsgBox sMsg, vbExclamation, ZapOnlineTitulo(TITULO)
    lblProgresso.Caption = "O envio por WhatsApp não faz parte do plano desta empresa."
    mForaDoPlano = True
    cmdEnviar.Enabled = False
ElseIf sSit <> "" And sSit <> "conectado" Then
    lblProgresso.Caption = "Atenção: o WhatsApp da empresa não está conectado (Sistema > WhatsApp)."
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And Not mEnviando And Not mCarregandoLista Then Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mCarregandoLista Then Cancel = 1: Exit Sub
If mEnviando Then
    MsgBox "Aguarde terminar o envio ou clique em Parar.", vbInformation, TITULO
    Cancel = 1
End If
End Sub

'---------------------------------------------------------------------------
' Lista de hoje
'---------------------------------------------------------------------------
'sDesmarcados = "|cod|cod|" dos clientes que o usuário tinha desmarcado (continuam desmarcados);
'lSelCliente = cliente que volta selecionado
Private Sub Carregar(Optional ByVal sDesmarcados As String = "", Optional ByVal lSelCliente As Long = 0)
Dim i As Long, r As Long, nProntos As Long, nSel As Long
Dim nFiltro As Integer, sSituacao As String

Me.MousePointer = vbHourglass
mCarregandoLista = True
HabilitarTela False
cmdParar.Enabled = False
Aguarde True
Cob_LerConfig mCfg
nFiltro = FiltroAtual()
Cob_Eventos = True     'DoEvents durante a leitura: o GIF continua girando
mQtd = Cob_Pendentes(Date, mItens, (chkJuros.Value = vbChecked), nFiltro)
Cob_Eventos = False
ReDim mMarca(0 To mQtd)

mSemEvento = True
With grdEnvio
    .Redraw = False
    .rows = 1
    .Cols = 8
    .TextMatrix(0, C_MARCA) = "Env."
    .TextMatrix(0, C_CLIENTE) = "Cliente"
    .TextMatrix(0, C_CELULAR) = "Celular"
    .TextMatrix(0, C_REGRA) = "Regra"
    .TextMatrix(0, C_PARC) = "Parc."
    .TextMatrix(0, C_TOTAL) = "Total"
    .TextMatrix(0, C_ATRASO) = IIf(nFiltro = COB_FILTRO_VENCIDO, "Atraso", "Vence")
    .TextMatrix(0, C_SITUACAO) = "Situação"
    .ColAlignment(C_MARCA) = flexAlignCenterCenter
    .ColAlignment(C_CLIENTE) = flexAlignLeftCenter
    .ColAlignment(C_CELULAR) = flexAlignLeftCenter
    .ColAlignment(C_REGRA) = flexAlignLeftCenter
    .ColAlignment(C_PARC) = flexAlignCenterCenter
    .ColAlignment(C_TOTAL) = flexAlignRightCenter
    .ColAlignment(C_ATRASO) = flexAlignCenterCenter
    .ColAlignment(C_SITUACAO) = flexAlignLeftCenter
    For i = 1 To mQtd
        .rows = .rows + 1
        r = .rows - 1
        .RowData(r) = i
        .TextMatrix(r, C_MARCA) = ""
        .TextMatrix(r, C_CLIENTE) = mItens(i).Nome
        .TextMatrix(r, C_CELULAR) = mItens(i).Celular
        .TextMatrix(r, C_REGRA) = mItens(i).Regra
        .TextMatrix(r, C_PARC) = mItens(i).QtdParcelas
        .TextMatrix(r, C_TOTAL) = Format$(mItens(i).Total, "#,##0.00")
        Select Case nFiltro
            Case COB_FILTRO_A_VENCER: .TextMatrix(r, C_ATRASO) = IIf(mItens(i).Vencimento > 0, Format$(mItens(i).Vencimento, "dd/mm/yyyy"), "-")
            Case COB_FILTRO_VENCENDO: .TextMatrix(r, C_ATRASO) = "Hoje"
            Case Else: .TextMatrix(r, C_ATRASO) = IIf(mItens(i).DiasAtraso > 0, mItens(i).DiasAtraso & "d", "-")
        End Select
        If Len(mItens(i).Motivo) = 0 Then
            .TextMatrix(r, C_SITUACAO) = "Pronto para enviar"
            nProntos = nProntos + 1
        Else
            .TextMatrix(r, C_SITUACAO) = mItens(i).Motivo
            PintarLinha r, &H808080
        End If
        MarcarLinha r, (InStr(sDesmarcados, "|" & mItens(i).CodCliente & "|") = 0)
        If lSelCliente <> 0 And mItens(i).CodCliente = lSelCliente And nSel = 0 Then nSel = r
    Next
    AjustarColunas grdEnvio, ",1,3,7,", C_MARCA
    If .rows > 1 Then
        If nSel = 0 Then nSel = 1
        .Row = nSel
        .Col = C_CLIENTE
        If nSel > 3 Then .TopRow = nSel - 3 Else .TopRow = 1
    End If
    .Redraw = True
End With
mSemEvento = False

Select Case nFiltro
    Case COB_FILTRO_A_VENCER: sSituacao = "com parcelas a vencer"
    Case COB_FILTRO_VENCENDO: sSituacao = "com parcelas vencendo hoje"
    Case Else: sSituacao = "com parcelas vencidas"
End Select
lblResumo.Caption = "Resumo: " & mQtd & " cliente(s) " & sSituacao & ", " & nProntos & " pronto(s) para enviar hoje.   " & _
                    "Enviadas hoje: " & Cob_EnviadosHoje() & " de " & mCfg.LimiteDia & " (limite por dia)."
MostrarPrevia
Aguarde False
Me.MousePointer = vbDefault
HabilitarTela True
mCarregandoLista = False
End Sub

'Recarrega a lista mantendo os desmarcados e a linha selecionada
Private Sub Recarregar()
Dim i As Long, lSel As Long

i = ItemDaLinha()
If i >= 1 Then lSel = mItens(i).CodCliente
Carregar DesmarcadosAtuais(), lSel
End Sub

Private Function DesmarcadosAtuais() As String
Dim r As Long, i As Long, s As String

s = "|"
For r = 1 To grdEnvio.rows - 1
    i = grdEnvio.RowData(r)
    If i >= 1 And i <= mQtd Then
        If Not mMarca(i) Then s = s & mItens(i).CodCliente & "|"
    End If
Next
DesmarcadosAtuais = s
End Function

Private Sub PintarLinha(ByVal nLinha As Long, ByVal nCor As Long)
Dim c As Integer, r As Long, cc As Integer, bAnt As Boolean

bAnt = mSemEvento
mSemEvento = True
With grdEnvio
    r = .Row: cc = .Col
    .Row = nLinha
    For c = 0 To .Cols - 1
        .Col = c
        .CellForeColor = nCor
    Next
    If r > 0 And r < .rows Then .Row = r
    .Col = cc
End With
mSemEvento = bAnt
End Sub

'Marca/desmarca a linha (imagem ImgMarcada/imgDesmarcada na coluna Env.). Quem não pode receber fica sem imagem.
Private Sub MarcarLinha(ByVal nLinha As Long, ByVal bMarcar As Boolean)
Dim i As Long, r As Long, cc As Integer, bAnt As Boolean

If nLinha < 1 Or nLinha >= grdEnvio.rows Then Exit Sub
i = grdEnvio.RowData(nLinha)
If i < 1 Or i > mQtd Then Exit Sub
If Len(mItens(i).Motivo) > 0 Then bMarcar = False
mMarca(i) = bMarcar

bAnt = mSemEvento
mSemEvento = True
With grdEnvio
    r = .Row: cc = .Col
    .Row = nLinha
    .Col = C_MARCA
    If Len(mItens(i).Motivo) > 0 Then
        Set .CellPicture = LoadPicture()
    ElseIf bMarcar Then
        Set .CellPicture = ImgMarcada.Picture
    Else
        Set .CellPicture = imgDesmarcada.Picture
    End If
    .CellPictureAlignment = flexAlignCenterCenter
    If r > 0 And r < .rows Then .Row = r
    .Col = cc
End With
mSemEvento = bAnt
End Sub

Private Function Marcada(ByVal nLinha As Long) As Boolean
Dim i As Long

If nLinha < 1 Or nLinha >= grdEnvio.rows Then Exit Function
i = grdEnvio.RowData(nLinha)
If i >= 1 And i <= mQtd Then Marcada = mMarca(i)
End Function

'Larguras pelo conteúdo: as colunas de sElasticas (",1,3,") encolhem ou crescem para o grid caber sem rolagem horizontal
Private Sub AjustarColunas(grd As MSFlexGrid, ByVal sElasticas As String, ByVal nColImagem As Integer)
Dim c As Long, r As Long, w() As Long, wt As Long, nTotal As Long, nDisp As Long, nElast As Long, nDif As Long

Me.FontName = grd.Font.Name
Me.FontSize = grd.Font.Size
Me.FontBold = False
ReDim w(0 To grd.Cols - 1)
For c = 0 To grd.Cols - 1
    If c = nColImagem Then
        w(c) = 450
    Else
        For r = 0 To grd.rows - 1
            wt = Me.TextWidth(grd.TextMatrix(r, c))
            If wt > w(c) Then w(c) = wt
        Next
        w(c) = w(c) + 180
    End If
    nTotal = nTotal + w(c)
    If InStr(sElasticas, "," & c & ",") > 0 Then nElast = nElast + w(c)
Next

nDisp = grd.Width - 90
If grd.rows * grd.RowHeight(0) > grd.Height - 60 Then nDisp = nDisp - 270   'barra de rolagem vertical
If nElast > 0 And nTotal <> nDisp Then
    nDif = nDisp - nTotal    'positivo = sobra espaço; negativo = falta
    For c = 0 To grd.Cols - 1
        If InStr(sElasticas, "," & c & ",") > 0 Then
            w(c) = w(c) + CLng(CDbl(nDif) * w(c) / nElast)
            If w(c) < 600 Then w(c) = 600
        End If
    Next
End If
For c = 0 To grd.Cols - 1
    grd.ColWidth(c) = w(c)
Next
End Sub

Private Function ItemDaLinha() As Long
If grdEnvio.Row < 1 Or grdEnvio.Row >= grdEnvio.rows Then Exit Function
ItemDaLinha = grdEnvio.RowData(grdEnvio.Row)
End Function

Private Sub MostrarPrevia()
Dim i As Long

i = ItemDaLinha()
If i < 1 Or i > mQtd Then
    txtPrevia.Text = ""
    Exit Sub
End If
txtPrevia.Text = "Para: " & IIf(Len(mItens(i).Celular) > 0, mItens(i).Celular, "(sem celular)") & vbCrLf & _
                 "Regra: " & IIf(Len(mItens(i).Regra) > 0, mItens(i).Regra, "(nenhuma)") & IIf(mItens(i).CodRegra = 0 And Len(mItens(i).Regra) > 0, " (prevista)", "") & vbCrLf & String$(30, "-") & vbCrLf & Cob_Texto(mItens(i), mCfg, True)
End Sub

Private Sub grdEnvio_Click()
If mEnviando Or mCarregandoLista Then Exit Sub
If grdEnvio.MouseCol = C_MARCA And grdEnvio.MouseRow = grdEnvio.Row Then
    MarcarLinha grdEnvio.Row, Not Marcada(grdEnvio.Row)
End If
MostrarPrevia
End Sub

Private Sub grdEnvio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 32 Or mEnviando Or mCarregandoLista Then Exit Sub   'espaço marca/desmarca
MarcarLinha grdEnvio.Row, Not Marcada(grdEnvio.Row)
KeyAscii = 0
End Sub

Private Sub grdEnvio_RowColChange()
If mSemEvento Or mEnviando Or mCarregandoLista Then Exit Sub
MostrarPrevia
End Sub

Private Sub cmdMarcarTodos_Click()
Marcar True
End Sub

Private Sub cmdDesmarcar_Click()
Marcar False
End Sub

Private Sub Marcar(ByVal bMarcar As Boolean)
Dim r As Long

grdEnvio.Redraw = False
For r = 1 To grdEnvio.rows - 1
    MarcarLinha r, bMarcar
Next
grdEnvio.Redraw = True
End Sub

Private Sub cmdNaoCobrar_Click()
Dim i As Long, bJaMarcado As Boolean

i = ItemDaLinha()
If i < 1 Then
    MsgBox "Clique na linha do cliente.", vbInformation, TITULO
    Exit Sub
End If
bJaMarcado = (Val(CStr(SQLExecutaRetorno("SELECT ISNULL(NaoCobrarWhatsApp, 0) AS n FROM cliente WHERE CODIGO = " & mItens(i).CodCliente, "n", 0))) <> 0)
If bJaMarcado Then
    If MsgBox("O cliente " & mItens(i).Nome & " está marcado para NÃO receber cobrança pelo WhatsApp." & vbCrLf & vbCrLf & _
              "Voltar a cobrar este cliente?", vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub
Else
    If MsgBox("Marcar o cliente " & mItens(i).Nome & " para NÃO receber cobrança pelo WhatsApp?" & vbCrLf & vbCrLf & _
              "Ele sai da régua até você desmarcar aqui.", vbQuestion + vbYesNo + vbDefaultButton2, TITULO) <> vbYes Then Exit Sub
End If
If Not Cob_NaoCobrar(mItens(i).CodCliente, Not bJaMarcado) Then
    MsgBox "Não foi possível gravar no cadastro do cliente.", vbExclamation, TITULO
    Exit Sub
End If
Recarregar
End Sub

'Abre o cadastro do cliente da linha; ao voltar, a lista é recarregada com os dados novos (celular, nome...)
Private Sub cmdAtualizarCliente_Click()
Dim i As Long

If mEnviando Then Exit Sub
i = ItemDaLinha()
If i < 1 Then
    MsgBox "Clique na linha do cliente que deseja atualizar.", vbInformation, TITULO
    Exit Sub
End If
Load Clientes_Cadastro
Clientes_Cadastro.EditarCliente mItens(i).CodCliente
Clientes_Cadastro.Show vbModal, Me
Unload Clientes_Cadastro
Recarregar
CarregarHistorico
End Sub

Private Sub cmdAtualizar_Click()
Recarregar
CarregarHistorico
End Sub

Private Sub cmdRegua_Click()
frmCobrancaRegua.Show vbModal
Recarregar
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cmdParar_Click()
Cob_Parar = True
cmdParar.Enabled = False
lblProgresso.Caption = "Parando depois da mensagem atual..."
End Sub

Private Sub cboPeriodo_Click()
If mIniciou Then CarregarHistorico
End Sub

'---------------------------------------------------------------------------
' Aguarde (GIF animado embutido no executável)
'---------------------------------------------------------------------------
'Grava o GIF do recurso na pasta temporária e monta a página que o WebBrowser mostra
Private Function AguardePreparar() As Boolean
Dim b() As Byte, sPasta As String, sGif As String, f As Integer
Dim nW As Long, nH As Long, nPx As Long

On Error GoTo Falha
sPasta = Environ$("TEMP")
If Len(sPasta) = 0 Then sPasta = App.path
If Right$(sPasta, 1) <> "\" Then sPasta = sPasta & "\"
sGif = sPasta & "oc_aguarde.gif"

b = LoadResData(101, "GIF")
On Error Resume Next
If Len(Dir$(sGif)) > 0 Then
    If FileLen(sGif) <> UBound(b) - LBound(b) + 1 Then Kill sGif
End If
On Error GoTo Falha
If Len(Dir$(sGif)) = 0 Then
    f = FreeFile
    Open sGif For Binary Access Write As #f
    Put #f, , b
    Close #f
End If

nW = WebBrowser1.Width \ Screen.TwipsPerPixelX
nH = WebBrowser1.Height \ Screen.TwipsPerPixelY
nPx = IIf(nW < nH, nW, nH) - 8
If nPx < 32 Then nPx = 32
mAguardeHtml = sPasta & "oc_aguarde.htm"
f = FreeFile
Open mAguardeHtml For Output As #f
Print #f, "<html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge""></head>"
Print #f, "<body scroll=""no"" style=""margin:0;padding:0;border:0;overflow:hidden;background:#FFFFFF;text-align:center"">"
Print #f, "<img src=""oc_aguarde.gif"" width=""" & nPx & """ height=""" & nPx & """ style=""margin-top:" & (nH - nPx) \ 2 & "px"">"
Print #f, "</body></html>"
Close #f
AguardePreparar = True
Exit Function
Falha:
mAguardeHtml = ""
Err.Clear
End Function

Private Sub Aguarde(ByVal bMostrar As Boolean)
Dim t As Single, bAcabouDeMostrar As Boolean

On Error Resume Next
If bMostrar Then
    If Len(mAguardeHtml) = 0 Then
        If Not AguardePreparar() Then Exit Sub
    End If
    If Not mAguardeCarregado Then
        WebBrowser1.Navigate mAguardeHtml
        mAguardeCarregado = True
    End If
    If Not WebBrowser1.Visible Then
        WebBrowser1.ZOrder 0
        WebBrowser1.Visible = True
        bAcabouDeMostrar = True
    End If
    'O Navigate é assíncrono: espera a página (e o GIF) carregar, até 3 s, senão o quadro aparece em branco
    t = Timer
    Do While WebBrowser1.ReadyState <> 4   'READYSTATE_COMPLETE
        DoEvents
        If Abs(Timer - t) > 3 Then Exit Do
    Loop
    If bAcabouDeMostrar Then
        'Dá tempo de o WebBrowser redesenhar o GIF antes da leitura (que segura a tela)
        t = Timer
        Do
            DoEvents
            Sleep 20
        Loop While Abs(Timer - t) < 0.4
    End If
Else
    WebBrowser1.Visible = False
End If
Err.Clear
End Sub

'---------------------------------------------------------------------------
' Envio dos marcados
'---------------------------------------------------------------------------
Private Sub cmdEnviar_Click()
Dim r As Long, i As Long, n As Long, nOk As Long, nErro As Long, nFeitos As Long
Dim nRestam As Long, nSeg As Long, sMsg As String, sNum As String, sNome As String, sSit As String
Dim marcados() As Long, bForaHorario As Boolean, sHora As String

'Marcados
ReDim marcados(0)
For r = 1 To grdEnvio.rows - 1
    If Marcada(r) Then
        n = n + 1
        ReDim Preserve marcados(n)
        marcados(n) = r
    End If
Next
If n = 0 Then
    MsgBox "Marque os clientes que vão receber a cobrança (coluna Env.).", vbInformation, TITULO
    Exit Sub
End If

'Limite por dia
Cob_LerConfig mCfg
nRestam = mCfg.LimiteDia - Cob_EnviadosHoje()
If nRestam <= 0 Then
    MsgBox "O limite de " & mCfg.LimiteDia & " mensagens por dia já foi atingido." & vbCrLf & vbCrLf & _
           "O limite protege o número da empresa de bloqueio pelo WhatsApp. Ele pode ser alterado na Régua de cobrança.", vbExclamation, TITULO
    Exit Sub
End If
If n > nRestam Then
    If MsgBox("Faltam só " & nRestam & " mensagens para o limite de hoje (" & mCfg.LimiteDia & ")." & vbCrLf & vbCrLf & _
              "Enviar para os " & nRestam & " primeiros marcados?", vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub
    n = nRestam
End If

'Horário e dia da semana
sHora = Format$(Now, "hh:mm")
bForaHorario = (sHora < Cob_Hora(mCfg.HoraInicio) Or sHora > Cob_Hora(mCfg.HoraFim))
If Weekday(Date) = vbSaturday And Not mCfg.EnviarSabado Then bForaHorario = True
If Weekday(Date) = vbSunday And Not mCfg.EnviarDomingo Then bForaHorario = True
If bForaHorario Then
    If MsgBox("Agora está fora do horário/dia de envio configurado (" & mCfg.HoraInicio & " às " & mCfg.HoraFim & ")." & vbCrLf & vbCrLf & _
              "Enviar mesmo assim?", vbQuestion + vbYesNo + vbDefaultButton2, TITULO) <> vbYes Then Exit Sub
End If

'WhatsApp conectado?
Me.MousePointer = vbHourglass
sSit = ZapOnlineSituacao(sNum, sNome, sMsg)
Me.MousePointer = vbDefault
If sSit = "" Then
    MsgBox sMsg, vbExclamation, ZapOnlineTitulo(TITULO)
    Exit Sub
ElseIf sSit <> "conectado" Then
    MsgBox "O WhatsApp da empresa não está conectado." & vbCrLf & vbCrLf & _
           "Conecte em Sistema > WhatsApp, lendo o QR code com o celular da empresa.", vbExclamation, TITULO
    Exit Sub
End If

If MsgBox("Enviar a cobrança para " & n & " cliente(s)?" & vbCrLf & vbCrLf & _
          "Entre uma mensagem e outra o sistema espera de " & mCfg.IntervaloMin & " a " & mCfg.IntervaloMax & " segundos " & _
          "(cerca de " & Int((n - 1) * (mCfg.IntervaloMin + mCfg.IntervaloMax) / 2 / 60 + 1) & " minuto(s) no total). " & _
          "Você pode clicar em Parar a qualquer momento.", vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub

'Envia
mEnviando = True
Cob_Parar = False
HabilitarTela False
Aguarde True
For r = 1 To n
    i = grdEnvio.RowData(marcados(r))
    grdEnvio.Row = marcados(r)
    grdEnvio.TopRow = IIf(marcados(r) > 3, marcados(r) - 3, 1)
    lblProgresso.Caption = "Enviando " & r & " de " & n & ": " & mItens(i).Nome & "..."
    grdEnvio.TextMatrix(marcados(r), C_SITUACAO) = "Enviando..."
    DoEvents

    If Cob_Enviar(mItens(i), mCfg, True, sMsg) Then
        nOk = nOk + 1
        grdEnvio.TextMatrix(marcados(r), C_SITUACAO) = "Enviada " & Format$(Now, "hh:mm")
        PintarLinha marcados(r), &H8000&
    Else
        nErro = nErro + 1
        grdEnvio.TextMatrix(marcados(r), C_SITUACAO) = "Erro: " & Replace(sMsg, vbCrLf, " ")
        PintarLinha marcados(r), &HC0&
        If ZapOnlineForaDoPlano() Then
            mForaDoPlano = True
            Aguarde False
            MsgBox sMsg, vbExclamation, ZapOnlineTitulo(TITULO)
            Exit For
        End If
    End If
    mItens(i).Motivo = "Já recebeu cobrança hoje"
    MarcarLinha marcados(r), False
    nFeitos = r

    If r < n Then
        nSeg = Cob_Intervalo(mCfg)
        lblProgresso.Caption = "Enviadas " & nOk & " de " & n & IIf(nErro > 0, " (" & nErro & " com erro)", "") & _
                               ". Próxima em " & nSeg & " segundos..."
        If Not Cob_Esperar(nSeg) Then Exit For
    End If
Next
Aguarde False
mEnviando = False
HabilitarTela True
lblProgresso.Caption = "Concluído: " & nOk & " enviada(s)" & IIf(nErro > 0, ", " & nErro & " com erro", "") & _
                       IIf(nFeitos < n, " (interrompido: " & n - nFeitos & " não enviada(s))", "") & "."
CarregarHistorico
lblResumo.Caption = "Resumo: enviadas hoje " & Cob_EnviadosHoje() & " de " & mCfg.LimiteDia & " (limite por dia). Clique em Atualizar para recarregar a lista."
End Sub

Private Sub HabilitarTela(ByVal bHabilitar As Boolean)
Dim k As Integer

'Antes de desabilitar, o foco vai para o grid: se ficasse num botão desabilitado, o VB passaria o foco para o
'próximo controle e, caindo num option (Vencido...), ele seria marcado sozinho e a lista trocaria de filtro
If Not bHabilitar Then
    On Error Resume Next
    grdEnvio.SetFocus
    Err.Clear
    On Error GoTo 0
End If
cmdEnviar.Enabled = bHabilitar And Not mForaDoPlano
cmdMarcarTodos.Enabled = bHabilitar
cmdDesmarcar.Enabled = bHabilitar
cmdNaoCobrar.Enabled = bHabilitar
cmdAtualizarCliente.Enabled = bHabilitar
cmdAtualizar.Enabled = bHabilitar
cmdRegua.Enabled = bHabilitar
cmdFechar.Enabled = bHabilitar
cboPeriodo.Enabled = bHabilitar
chkJuros.Enabled = bHabilitar
For k = COB_FILTRO_A_VENCER To COB_FILTRO_VENCIDO
    optFiltro(k).Enabled = bHabilitar
Next
cmdParar.Enabled = Not bHabilitar
End Sub

'---------------------------------------------------------------------------
' Histórico
'---------------------------------------------------------------------------
Private Sub CarregarHistorico()
Dim rs As ADODB.Recordset, dIni As Date, c As Integer

Select Case cboPeriodo.ListIndex
    Case 1: dIni = Date - 6
    Case 2: dIni = Date - 29
    Case Else: dIni = Date
End Select

With grdHist
    .Redraw = False
    .rows = 1
    .Cols = 8
    .TextMatrix(0, 0) = "Data/hora"
    .TextMatrix(0, 1) = "Cliente"
    .TextMatrix(0, 2) = "Celular"
    .TextMatrix(0, 3) = "Regra"
    .TextMatrix(0, 4) = "Total"
    .TextMatrix(0, 5) = "Situação"
    .TextMatrix(0, 6) = "Origem"
    .TextMatrix(0, 7) = "Detalhe"
    For c = 0 To 7
        .ColAlignment(c) = flexAlignLeftCenter
    Next
    .ColAlignment(4) = flexAlignRightCenter
End With

On Error GoTo Falha
Set rs = dbData.OpenRecordset("SELECT e.DataHora, c.NOME, e.Celular, r.Descricao, e.ValorTotal, e.Situacao, e.Erro, e.Manual, e.Usuario " & _
    "FROM CobrancaEnvio e LEFT JOIN cliente c ON c.CODIGO = e.CodCliente LEFT JOIN CobrancaRegra r ON r.Id = e.CodRegra " & _
    "WHERE e.DataHora >= '" & Format$(dIni, "yyyymmdd") & "' ORDER BY e.DataHora DESC")
Do While Not rs.EOF
    grdHist.rows = grdHist.rows + 1
    grdHist.TextMatrix(grdHist.rows - 1, 0) = Format$(rs("DataHora"), "dd/mm/yyyy hh:mm")
    grdHist.TextMatrix(grdHist.rows - 1, 1) = "" & rs("NOME")
    grdHist.TextMatrix(grdHist.rows - 1, 2) = "" & rs("Celular")
    grdHist.TextMatrix(grdHist.rows - 1, 3) = IIf(IsNull(rs("Descricao")), "(regra excluída)", "" & rs("Descricao"))
    grdHist.TextMatrix(grdHist.rows - 1, 4) = Format$(rs("ValorTotal"), "#,##0.00")
    grdHist.TextMatrix(grdHist.rows - 1, 5) = IIf(rs("Situacao") = 1, "Enviada", "Erro")
    grdHist.TextMatrix(grdHist.rows - 1, 6) = IIf(CBool(rs("Manual")), "Manual" & IIf(Len("" & rs("Usuario")) > 0, " (" & rs("Usuario") & ")", ""), "Automático")
    grdHist.TextMatrix(grdHist.rows - 1, 7) = Replace("" & rs("Erro"), vbCrLf, " ")
    rs.MoveNext
Loop
rs.Close

continua:
On Error Resume Next
AjustarColunas grdHist, ",1,3,7,", -1
grdHist.Redraw = True
Err.Clear
Exit Sub
Falha:
Resume continua
End Sub
