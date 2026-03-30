VERSION 5.00
Begin VB.Form Exportar_XML 
   ClientHeight    =   7020
   ClientLeft      =   7995
   ClientTop       =   3450
   ClientWidth     =   12015
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "pExpXML"
   LockControls    =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   12015
   Begin VB.PictureBox Painel 
      Height          =   3270
      Index           =   0
      Left            =   15
      ScaleHeight     =   3210
      ScaleWidth      =   5670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   5730
      Begin VB.CheckBox ChkCp 
         Caption         =   "Incluir Entradas?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   3
         Left            =   4080
         TabIndex        =   18
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Incluir PDF?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   2
         Left            =   2820
         TabIndex        =   17
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Enviar Email?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   1200
         Width           =   1260
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Compactar?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1200
         Width           =   1200
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1845
         TabIndex        =   14
         Top             =   870
         Width           =   3465
      End
      Begin VB.CommandButton bottxtCampo3 
         Height          =   195
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   870
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   870
         TabIndex        =   12
         Top             =   870
         Width           =   480
      End
      Begin VB.CommandButton bottxtCampo2 
         Height          =   195
         Left            =   1365
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   870
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Height          =   255
         Index           =   1
         Left            =   195
         MaxLength       =   2
         TabIndex        =   10
         Top             =   840
         Width           =   450
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   225
         MaxLength       =   254
         TabIndex        =   9
         Top             =   1710
         Width           =   5085
      End
      Begin VB.CommandButton bottxtCampo0 
         Height          =   195
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1710
         Width           =   255
      End
      Begin VB.PictureBox Painel 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   195
         ScaleHeight     =   360
         ScaleWidth      =   5415
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2235
         Width           =   5415
      End
      Begin VB.PictureBox botCancela 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         ScaleHeight     =   360
         ScaleWidth      =   1335
         TabIndex        =   6
         Top             =   2700
         Width           =   1400
      End
      Begin VB.PictureBox botOk 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1140
         ScaleHeight     =   360
         ScaleWidth      =   1335
         TabIndex        =   5
         Top             =   2700
         Width           =   1400
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informe a Empresa/Filial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1815
         TabIndex        =   28
         Top             =   640
         Width           =   1725
      End
      Begin VB.Label labFdo3 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1815
         TabIndex        =   27
         Top             =   840
         Width           =   3795
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   26
         Top             =   640
         Width           =   285
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mês"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   25
         Top             =   640
         Width           =   285
      End
      Begin VB.Label labFdo2 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informe o diretório de destino dos arquivos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   23
         Top             =   1480
         Width           =   3105
      End
      Begin VB.Label labFdo0 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   195
         TabIndex        =   22
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   4890
         TabIndex        =   21
         Top             =   1995
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   20
         Top             =   1995
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exportar Arquivos XML"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   935
         TabIndex        =   19
         Top             =   135
         Width           =   4005
      End
   End
   Begin VB.ComboBox cboPeriodo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1515
   End
   Begin VB.ComboBox cboAno 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   675
   End
   Begin VB.Label lblAno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ano"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   345
   End
End
Attribute VB_Name = "Exportar_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* Sistema...: SISC
'* Empresa...: RONDOSOFT SYSTEMS
'* Módulo....: pExpXML
'* Função....: Exportar XML
'* CopyRight.: (C)2022 RONDOSOFT SYSTEMS
'* Criação...: GAS-2003, GERADOR AUTOMÁTICO DE SISTEMAS
'* Data......: 06/04/2022 15:01:35
'* * * * * * *

Option Explicit
DefInt A-Z

Public vgSituacao As Integer                      'situação de edição que do módulo
Public vgCaracteristica As Integer                'caracteristica do módulo
Public vgTipo As Integer                          'tipo do módulo
Public vgPriVez As Integer                        'flag de carregamento do módulo
Public vgFormID As Long                           'identificador único para o módulo
Public vgIdentTab As String                       'nome da tabela principal do módulo
Public vgEmBrowse As Integer                      'flag se o módulo esta em grade
Public vgTemFiltro As Integer                     'flag se tem ou não filtro no módulo
Public vgCancelou As Integer                      'flag cancela ou nao o processo do módulo
Public vgQueryAtual As String                     'expressão SQL definada para o módulo

'Public vgFP As New frmPreview                     'formuário de preview do relatório
'Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
'Public vgFG As frmGauge                           'formulário do gauge
Public vgQReg As Long                             'quantidade de já feira para o gauge

Dim txtCampo(3) As New FormataCampos              'classe dos compos tipo texto do módulo
Dim chkCampo(3) As New FormataCampos              'classe dos compos tipo lógico do módulo
Dim DiretorioDestino As String, xMes As String, xAno As String
Dim IdFilial As Integer, ckCompactar As Boolean, ckEnviarEmail As Boolean
Dim ckPDF As Boolean, ckEntrada As Boolean
Public btExportar As Object, lblStatus As Object, frameBarra As Object
Public xRegistros As Object, cpDiretorioDestino As Object, cpMes As Object
Public cpAno As Object
Dim TbFiliais As New ADODB.Recordset, TbNFe As New ADODB.Recordset, TbNFCe As New ADODB.Recordset
Dim TbEntradaEstoque As New ADODB.Recordset

Public WithEvents Prog As ComctlLib.ProgressBar
Attribute Prog.VB_VarHelpID = -1
Dim sTemp As String, Tipo As String
Dim Tm As Long, PReg As Long, RegS As Long
Dim NArq As Byte
Dim Sair As Boolean
Dim xCNPJ As String

'Retorna Ação FormLoad
Private Sub FormLoad_F9()
  CentraNaTela Me
  'Sair = True
   
  'adiciona a barra de progresso
  Set Prog = Me.Controls.Add("COMCTL.ProgCtrl.1", "BarraProg1", frameBarra)
  'Prog.Scrolling = ccScrollingSmooth
  '=========Esquerda, Top, Largura, Altura
  Prog.Move 0, 0, frameBarra.Width, frameBarra.Height
  Prog.Visible = True
  Prog.Max = 100
  xRegistros.Caption = 0
  Mostra "Preencha os campos!"
End Sub

Private Sub Mostra(Texto As String)
  lblStatus.Caption = Texto
  DoEvents
End Sub

Private Sub Ini_Proc(Texto As String, REG As Long)
  Prog.Max = REG
  lblStatus.Caption = Texto
  DoEvents
End Sub

Private Sub Inc_proc()
  If Prog.Value < Prog.Max Then
    Prog.Value = Prog.Value + 1
    xRegistros.Caption = Prog.Value & " de " & Prog.Max
    DoEvents
  End If
End Sub

Private Sub PegaDiretorio_F9(Campo As Object)
Dim bi As BrowseInfo, vgRet As Variant, vgDirCapt As String
  bi.hOwner = Me.hWnd
  bi.lpszTitle = "Selecione o Diretório onde deverão ficar os arquivos XML."
  bi.ulFlags = BI_RETORNADIR Or BIF_EDITBOX Or BIF_USENEWUI
  vgRet = SHBrowseForFolder(bi)                                     'vamos chamar o diálog
  vgDirCapt$ = Space$(512)                                          'prepara o buffer
  If SHGetPathFromIDList(ByVal vgRet, ByVal vgDirCapt$) Then        'pega a pasta capturada e se foi cancelado ou não
    vgDirCapt$ = Left(vgDirCapt$, InStr(vgDirCapt$, Chr$(0)) - 1)    'pega o caminho da pasta capturada
    Campo.Text = vgDirCapt$ & "\"
  End If
  Campo.SetFocus
End Sub

Private Function PeriodoFiltro_F9() As String
  PeriodoFiltro_F9$ = "(Month(DataEmissao) = " & xMes & " AND Year(DataEmissao) = " & xAno & ")"
End Function

Private Function FiltroInicialNFe_F9() As String  ' {* Apelidos *} 1635
Dim FI As String
  cf FI, "NFeEnviada <> 0"
  cf FI, "IdFilialNF = " & IdFilial
  cf FI, PeriodoFiltro_F9
  FiltroInicialNFe_F9$ = FI
End Function

Private Function FiltroInicialNFCe_F9() As String
Dim FI As String
  cf FI, "NFeEnviada <> 0"
  cf FI, "IdFilialNF = " & IdFilial
  cf FI, PeriodoFiltro_F9
  FiltroInicialNFCe_F9$ = FI
End Function

Private Function FiltroInicialEntradas_F9() As String
Dim FI As String
  cf FI, "TipoEntradaEst = 'Nota Fiscal'"
  cf FI, "IdFilial = " & IdFilial
  cf FI, "Lancada <> 0"
  cf FI, "EmissaoPropria = 0"
  cf FI, "(Month(DtEntradaEst) = " & xMes & " AND Year(DtEntradaEst) = " & xAno & ")"
  FiltroInicialEntradas_F9$ = FI
End Function

'Retorna condição p/ processar campo FANTASIA, tabela TBEMPRESA
Private Function ExportarXMLNFe_F9() As Boolean  ' {* Apelidos *} 1635
Dim Xarquivo As String, xArquivoDestino As String, xArquivoCancelada As String, contXML As Long
Dim xDiretorioDestino As String
  If Vazio(xCNPJ) Then
    xCNPJ = TbFiliais!CNPJ
    xCNPJ = Retira(xCNPJ, "./- ", UM_A_UM)
  End If
  If Vazio(DiretorioDestino) Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
  End If
  If ckEnviarEmail Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
  End If
  Mostra "Exportando NFe nº " & LPad(TbNFe!NumeNota, 6, "0")
  Prog.Max = TbNFe.RecordCount
  Inc_proc
  contXML = 0
Volta:

  If contXML = 3 Then Exit Function

  dirXML = TbFiliais!DiretorioXML
  If Vazio(dirXML) Then dirXML = GetString(&H80000001, "nfe", "PathPrincipal")
  dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")

  Xarquivo = dirXML & "nfe\arquivos\procNFe\" & TbNFe!NFeChaveAcesso & "-procNFe.xml"
  xDiretorioDestino = DiretorioDestino & IIf(Right(DiretorioDestino, 1) = "\", "", "\") & "NFe\"
  If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino

  If Existe(Xarquivo) = -1 Then
    xArquivoDestino = xDiretorioDestino & TbNFe!NFeChaveAcesso & "-procNFe.xml"
    FileCopy Xarquivo, xArquivoDestino
    If ckPDF Then
      Dim sistNFe As snfe.Util
      Set sistNFe = New snfe.Util
      If Not Existe(dirXML & "nfe\arquivos\PDF\" & TbNFe!NFeChaveAcesso & ".pdf") Then
        iRetorno = ConfiguraDLLNFeNFCe(IdFilial, 55, "1", sistNFe)
        Call sistNFe.DANFeImprimir(Xarquivo, False, "", True, dirXML & "nfe\arquivos\PDF\" & TbNFe!NFeChaveAcesso & ".pdf", 0, TbNFe!NFeCancelada, True, "")
      End If
      xDiretorioDestino = DiretorioDestino & IIf(Right(DiretorioDestino, 1) = "\", "", "\") & "PDF\"
      If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino
        Xarquivo = dirXML & "nfe\arquivos\PDF\" & TbNFe!NFeChaveAcesso & ".pdf"
        xArquivoDestino = xDiretorioDestino & TbNFe!NFeChaveAcesso & ".pdf"
        FileCopy Xarquivo, xArquivoDestino
      End If
      If TbNFe!NFeCancelada Then
        contXML = 0
VoltaCancelada:
        If contXML = 3 Then Exit Function
        xArquivoCancelada = dirXML & "nfe\arquivos\procEventoNFe\env110111-" & TbNFe!NFeChaveAcesso & "-procEventoNFe.xml"
        If Existe(xArquivoCancelada) = -1 Then
          Xarquivo = xArquivoCancelada
          xArquivoDestino = xDiretorioDestino & TbNFe!NFeChaveAcesso & "-cancNFe.xml"
          FileCopy Xarquivo, xArquivoDestino
        Else
          'GeraArqCancelada TbNFe!CanceladaProtocolo
          'contXML = contXML + 1
          'GoTo VoltaCancelada
        End If
      End If
   Else
      NFeXML = SQLExecutaRetorno("SELECT XMLAutorizado FROM TbNotaFiscalProd WHERE NFeChaveAcesso = '" & TbNFe!NFeChaveAcesso & "'", "XMLAutorizado", "")
      If Vazio(NFeXML) Then NFeXML = SQLExecutaRetorno("SELECT XMLAutorizado FROM TbNotaFiscalProd_XML WHERE IDNFProd = '" & TbNFe!IdNFProd & "'", "XMLAutorizado", "")
      If Vazio(NFeXML) Then
        GeraArqDistribuicao TbNFe!NFeChaveAcesso, TbNFe!NFeRecibo
      Else
        xArquivoDestino = TbNFe!NFeChaveAcesso & "-procNFe.xml"
        xCaminhoXML = xDiretorioDestino
        fsoSalva xArquivoDestino, UTF8_Encode(NFeXML), xCaminhoXML
      End If
      contXML = contXML + 1
      GoTo Volta
   End If
   If Prog.Value = Prog.Max Then Prog.Value = 0
   ExportarXMLNFe_F9 = False  'Deve retornar um valor lógico
End Function

'Retorna condição p/ processar campo FANTASIA, tabela TBEMPRESA
Private Function ExportarXMLNFCe_F9() As Boolean  ' {* Apelidos *} 1636
Dim Xarquivo As String, xArquivoDestino As String, xArquivoCancelada As String, contXML As Long
Dim xDiretorioDestino As String
 If Vazio(xCNPJ) Then
    xCNPJ = TbFiliais!CNPJ
    xCNPJ = Retira(xCNPJ, "./- ", UM_A_UM)
 End If
 If Vazio(DiretorioDestino) Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
 End If
 If ckEnviarEmail Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
 End If
 Mostra "Exportando NFCe nº " & LPad(TbNFCe!NumeNota, 6, "0")
 If Prog.Max <> TbNFCe.RecordCount Then Prog.Value = 0
 Prog.Max = TbNFCe.RecordCount
 Inc_proc
 contXML = 0
Volta:

 If contXML = 3 Then Exit Function

 dirXML = TbFiliais!DiretorioXML
 If Vazio(dirXML) Then dirXML = GetString(&H80000001, "nfe", "PathPrincipal")
 dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")

 Xarquivo = dirXML & "nfe\arquivos\procNFe\" & TbNFCe!NFeChaveAcesso & "-procNFe.xml"
 xDiretorioDestino = DiretorioDestino & IIf(Right(DiretorioDestino, 1) = "\", "", "\") & "NFCe\"
 If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino

 If Existe(Xarquivo) = -1 Then
    xArquivoDestino = xDiretorioDestino & TbNFCe!NFeChaveAcesso & "-procNFe.xml"
    FileCopy Xarquivo, xArquivoDestino
    If TbNFCe!NFeCancelada Then
       contXML = 0
VoltaCancelada:
       If contXML = 3 Then Exit Function
       xArquivoCancelada = dirXML & "nfe\arquivos\procEventoNFe\env110111-" & TbNFCe!NFeChaveAcesso & "-procEventoNFe.xml"
       If Existe(xArquivoCancelada) = -1 Then
          Xarquivo = xArquivoCancelada
          xArquivoDestino = xDiretorioDestino & TbNFCe!NFeChaveAcesso & "-cancNFe.xml"
          FileCopy Xarquivo, xArquivoDestino
       Else
          'GeraArqCancelada TbNFCe!CanceladaProtocolo
          'contXML = contXML + 1
          'GoTo VoltaCancelada
       End If
    End If
 Else
    NFeXML = SQLExecutaRetorno("SELECT XMLAutorizado FROM TbNFCe WHERE NFeChaveAcesso = '" & TbNFCe!NFeChaveAcesso & "'", "XMLAutorizado", "")
    If Vazio(NFeXML) Then NFeXML = SQLExecutaRetorno("SELECT XMLAutorizado FROM TbNFCe_XML WHERE IDNFProd = '" & TbNFCe!IdNFProd & "'", "XMLAutorizado", "")
    If Vazio(NFeXML) Then
       GeraArqDistribuicao TbNFCe!NFeChaveAcesso, TbNFCe!NFeRecibo
    Else
       xArquivoDestino = TbNFCe!NFeChaveAcesso & "-procNFe.xml"
       xCaminhoXML = xDiretorioDestino
       fsoSalva xArquivoDestino, UTF8_Encode(NFeXML), xCaminhoXML
    End If
    contXML = contXML + 1
    GoTo Volta
 End If
 
 ExportarXMLNFCe_F9 = False  'Deve retornar um valor lógico
End Function

'Trata Ação do botão
Private Sub AposExportar_F9()
Dim FileZipName As String, PathToCompress As String, DestPath As String, FullPathZip As String
Dim NomeEmp As String, emailDestino As String, i As Integer
  If ckCompactar Then
     'btExportar.Enabled = False
     If IniciaComponenteCompactacao Then
     
        i = 0
        
        'Caminho para comprimir arquivo
        DestPath = DiretorioDestino
        
        RsOpen TbFiliais, "SELECT TOP 1 * FROM TbFiliais"
        Ptab TbFiliais, "1 Filiais", IdFilial
        'Nome a dar ao ficheiro se for Compactador Zip a sua extenção será .Zip / Se o compactador for Winrar a extensão será Rar
        '=========================================================================================================================
        'Neste caso escolhemos rar porque não nos oferece problemas de compactação de ficheiros que estão a ser utilizados pelo sistema.
        NomeEmp = TbFiliais!RazaoSocial
        NomeEmp = RemoveAcento(NomeEmp)
        NomeEmp = Substitui(NomeEmp, ".,/", "", UM_A_UM)
        NomeEmp = Substitui(NomeEmp, " ", "_", UM_A_UM)
        FileZipName = NomeEmp & "_" & Retira(TbFiliais!CNPJ, "./- ", UM_A_UM) & ".rar"
        
        'Caminho + ficheiro.rar
        FullPathZip = Transforma(App.path & IIf(Right(App.path, 1) = "\", "", "\") & FileZipName)
        
        'Caminho a comprimir (nota : no final não deixar no fim da directoria o simbolo de '\'
        DiretorioDestino = IIf(Right(DiretorioDestino, 1) = "\", Left(DiretorioDestino, Len(DiretorioDestino) - 1), DiretorioDestino)
        PathToCompress = Transforma(DiretorioDestino)
        
        'Chama o compressor que se encontra instalado para o efeito.
        If xWinRar <> "" Then
            Shell xWinRar & " a " & FullPathZip & " " & PathToCompress, vbNormalFocus ', vbHide
        Else
            Shell xWinZip & " -a " & FullPathZip & " " & PathToCompress, vbNormalFocus ', vbHide
        End If
        
        DoEvents
        
        Load frmECFMsg
        frmECFMsg.SetaMensagem "Aguarde! Compactando arquivos XML..."
        
        If ckEnviarEmail Then
           Do While Not (Existe(FullPathZip) = -1)
              DoEvents
              i = i + 1
              If i > 50 Then Exit Do
           Loop
           If Not (Existe(FullPathZip) = -1) Then Exit Sub
           emailDestino = InputBox("Informe o Email do destinatario", "Envio de Email", TbContabilista!Email)
           If Not Vazio(emailDestino) Then
              HourglassShow
              EnviaEmail_F9 emailDestino, FullPathZip
              DoEvents
              HourglassClose
           End If
        End If
     End If
     'btExportar.Enabled = True
  End If
End Sub

Private Function IniciaComponenteCompactacao() As Boolean
    On Error Resume Next
    'Verifica se está instalado um dos 2 compressores referênciados
    
    Dim MiObjeto As Object, Ret As String
    Set MiObjeto = CreateObject("Wscript.Shell")
    
    Ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinRAR\shell\open\command\")
    xWinRar = Left(Ret, InStrRev(Ret, " "))
    
    Ret = MiObjeto.RegRead("HKEY_CLASSES_ROOT\WinZip\shell\open\command\")
    xWinZip = Left(Ret, InStrRev(Ret, " "))
    
    IniciaComponenteCompactacao = True
    Set MiObjeto = Nothing
    
    If xWinRar & xWinZip = "" Then
       MsgBox "Não se encontra instalado o WinZip nem WinRar :("
       IniciaComponenteCompactacao = False
    End If
End Function

Private Sub EnviaEmail_F9(EmailPara As String, Anexo1 As String)
Dim emailDest As String, pathAnexo() As String, NomeRemetente As String, corpoEmail As String, emailCC() As String
Dim temParcelas As Boolean
Dim sistNFe As snfe.Util

    On Error GoTo DeuErro
    
    Set sistNFe = New snfe.Util
    
    emailDest = EmailPara
    
    If Vazio(emailDest) Then Exit Sub
    
    iRetorno = ConfiguraDLLNFeNFCe(IdFilial, 55, "1", sistNFe, False)

    ReDim pathAnexo(0)
    NomeRemetente = SQLExecutaRetorno("SELECT Fantasia FROM TbFiliais WHERE IdFilial = " & IdFilial, "Fantasia", TbEmpresa!Fantasia)
    pathAnexo(0) = Anexo1
    corpoEmail = "Segue em anexo os arquivos XML de NFe e NFCe emitidos no período. " & _
                 "<br><br>" & _
                 "Atenciosamente, " & _
                 "<br><br>" & _
                 "#nome_emitente#"
    corpoEmail = Substitui(corpoEmail, "#nome_emitente#", SQLExecutaRetorno("SELECT RazaoSocial FROM TbFiliais WHERE IdFilial = " & IdFilial, "RazaoSocial"), SO_UM)

    If (emailDest <> Empty) Then
       Screen.MousePointer = vbHourglass
       GeraLogAcao LoadGasString(295) + " Enviando Email XML para Contabilidade"
       iRetorno = sistNFe.EmailEnviar(emailDest, "Arquivos XML ref. Mês " & LPad(xMes, 2, "0") & "/" & LPad(xAno, 4, "0"), corpoEmail, pathAnexo, emailCC)
       Screen.MousePointer = vbDefault
    End If
    
    If iRetorno Then MsgBox "Email enviado com sucesso!", vbInformation + vbOKOnly, "EMAIL OK!"
    
    Set sistNFe = Nothing
    
    Exit Sub
    
DeuErro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "ERRO: Envio Email"
    Err.Clear
    Set sistNFe = Nothing
End Sub

'Retorna condição p/ processar campo RAZAOSOCIAL, tabela TBEMPRESA
Private Function ExportarXMLEntradas_F9() As Boolean  ' {* Apelidos *} 1925
Dim Xarquivo As String, xArquivoDestino As String, xArquivoCancelada As String, contXML As Long
Dim xDiretorioDestino As String
  If Vazio(xCNPJ) Then
    xCNPJ = TbFiliais!CNPJ
    xCNPJ = Retira(xCNPJ, "./- ", UM_A_UM)
  End If
  If Vazio(DiretorioDestino) Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
  End If
  If ckEnviarEmail Then
    DiretorioDestino = App.path & "\ExportarXML"
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xCNPJ
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
    DiretorioDestino = DiretorioDestino & "\" & xAno & xMes
    If Not Existe(DiretorioDestino) Then MkDir DiretorioDestino
  End If
  Mostra "Exportando NFe Entrada nº " & TbEntradaEstoque!NumeNota
  If Prog.Max <> TbEntradaEstoque.RecordCount Then Prog.Value = 0
  Prog.Max = TbEntradaEstoque.RecordCount
  Inc_proc
  contXML = 0
Volta:

  If contXML = 3 Then Exit Function

  dirXML = TbFiliais!DiretorioXML
  dirXML = IIf(Right(dirXML, 1) = "\", dirXML, dirXML & "\")

  Xarquivo = dirXML & "nfe\arquivos\ConfRecebto\" & TbEntradaEstoque!NFeChaveAcesso & "-procNFe.xml"
  xDiretorioDestino = DiretorioDestino & IIf(Right(DiretorioDestino, 1) = "\", "", "\") & "Entradas\"
  If Not Existe(xDiretorioDestino) Then MkDir xDiretorioDestino

  NFeXML = SQLExecutaRetorno("SELECT XMLAutorizado FROM TbEntradaEstoqueXML WHERE IdEntradaEstoque = " & TbEntradaEstoque!IDEntradaEst, "XMLAutorizado", "")
  If Not Vazio(NFeXML) Then
    xArquivoDestino = TbEntradaEstoque!NFeChaveAcesso & "-procNFe.xml"
    xCaminhoXML = xDiretorioDestino
    fsoSalva xArquivoDestino, UTF8_Encode(NFeXML), xCaminhoXML
  End If
  contXML = contXML + 1
  'GoTo Volta
  ExportarXMLEntradas_F9 = False  'Deve retornar um valor lógico
End Function

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
    On Error Resume Next                          'prepara para possíveis erros
    DiretorioDestino = txtCampo(0).Value
    xMes = txtCampo(1).Value
    xAno = txtCampo(2).Value
    IdFilial = txtCampo(3).Value
    ckCompactar = chkCampo(0).Value
    ckEnviarEmail = chkCampo(1).Value
    ckPDF = chkCampo(2).Value
    ckEntrada = chkCampo(3).Value
    If Err Then Err.Clear                         'se deu algum erro, vamos resetá-lo
End Sub

'poe filtro na lista externa (combobox)
Private Sub PoeFiltroCbo(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 3
            txtCampo(3).Filter = "TipoFilial = " & 0 & ""
    End Select
End Sub

'abre a tabela relacionada a lista externa (combobox)
Public Sub AbreTabelaCbo(Index As Integer)
    Dim x As String
    On Error Resume Next
    Select Case Index
        Case 3
            x$ = "SELECT Fantasia, CNPJ, IdFilial" + _
         " FROM TbFiliais" + _
         " WHERE TipoFilial = " & 0 & "" + _
         " ORDER BY Fantasia"
            If Err = 0 And x$ <> txtCampo(3).RecordSource.Source Then
                Screen.MousePointer = vbHourglass
                RsOpen txtCampo(3).RecordSource, x$
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub

'executa processos/validacoes nos campos do formuário
Public Function Executar(vgOq As String) As String
    Dim vgMsg As String, vgOk As Integer, vgNVez As Integer
    On Error GoTo DeuErro                         'fica na espera de um erro...
    vgMsg$ = ""                                   'retorna uma msg dizendo o motivo
    vgOk = True                                   'se a validação esta OK
    vgNVez = 0                                    'porque não fez o processo/validacoes
    If vgOq = VALIDACOES Then
        InicializaApelidos COM_TEXTBOX
        Ptab TbFiliais, "1 Filiais", IdFilial
        If vgOk Then
            vgOk = ((xMes >= "01") And xMes <= "12")
            vgMsg$ = "Mês inválido!"
            If Not vgOk Then txtCampo(1).SetFocus
        End If
        If vgOk Then
            vgOk = (IdFilial > 0)
            vgMsg$ = "Empresa/Filial inválido!"
            If Not vgOk Then txtCampo(3).SetFocus
        End If
        If vgOk Then vgMsg$ = ""
        DoEvents
    ElseIf vgOq = INICIALIZACOES Then
        InicializaApelidos COM_TEXTBOX
        txtCampo(1).Value = Format(DateAdd("m", -1, Date), "mm")
        txtCampo(2).Value = Format(DateAdd("m", -1, Date), "yyyy")
        txtCampo(3).Value = vgFilial
        chkCampo(0).Value = False
        chkCampo(1).Value = False
        chkCampo(2).Value = False
        chkCampo(3).Value = False
    ElseIf vgOq = INI_APELIDOS Then
        InicializaApelidos COM_REGISTRO
        ExecutaPreValidacao
    End If
    Executar = vgMsg$                             'prepara saida da função
    Exit Function                                 'e cai fora...

DeuErro:
    Select Case Err                               'vamos verificar se deu algum erro
        Case 3197                                 'registro acabou de ser alterado por outro usuário
            Resume
        Case 3046, 3158, 3186, 3187, 3188, 3189, 3218, 3260
            vgNVez = vgNVez + 1                   'outro usuário acessando
            If vgNVez <= 10 Then                  'vamos tentar 10 vezes
                Delay 0.5                         'mais antes espera 1/2 segundo
                Resume                            'antes de tentar
            End If
    End Select
    Executar = Err.Source + "|" + Trim$(Str$(Err)) + " - " + Error$ 'não teve jeito o erro não pode ser evitado...
End Function

'reposiciona arquivos relacionados
Private Sub ExecutaPTabs()
    Ptab TbFiliais, "1 Filiais", IdFilial
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
    Dim Ok As Boolean
    On Error Resume Next                          'prepara para possiveis erros
    Ok = (Not ckEnviarEmail)
    txtCampo(0).Enabled = Ok
    Ok = (Not ckEnviarEmail)
    Label(3).Enabled = Ok
    Ok = (ckCompactar)
    chkCampo(1).Enabled = Ok
    If Err Then Err.Clear                         'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
    If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
    InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
    txtCampo(Index).Change
    If Index = 2 Or Index = 3 Then
        ExecutaPTabs
        ExecutaPreValidacao
    End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
        On Error Resume Next
        'Select Case Index
        '    Case 1
        '        If ValBrasil(txtCp(1).Text) = 0 Then
        '            txtCampo(1).Value = Format(DateAdd("m", -1, Date), "mm")
        '            txtCp_Change Index
        '            InicializaApelidos COM_TEXTBOX
        '            ExecutaPTabs
        '            ExecutaPreValidacao
        '        End If
        '    Case 2
        '        If Len(txtCp(2).Text) = 0 Then
        '            txtCampo(2).Value = Format(DateAdd("m", -1, Date), "yyyy")
        '            txtCp_Change Index
        '            InicializaApelidos COM_TEXTBOX
        '            ExecutaPTabs
        '            ExecutaPreValidacao
        '        End If
        '    Case 3
        '        PoeFiltroCbo 3
        '        If Len(txtCp(3).Text) = 0 Then
        '            txtCampo(3).Value = vgFilial
        '            txtCp_Change Index
        '            InicializaApelidos COM_TEXTBOX
        '            ExecutaPTabs
        '            ExecutaPreValidacao
        '        End If
        'End Select
    txtCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    txtCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)
    txtCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
    txtCampo(Index).LostFocus                     'executa LostFocus na classe
    InicializaApelidos COM_TEXTBOX                'inicializa apelidos dos campos
    ExecutaPreValidacao                           'habilita/desabilita campos
End Sub

'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
    If vgPriVez Then Exit Sub
    InicializaApelidos COM_TEXTBOX
    ExecutaPreValidacao                           'habilita/desabilita campos
    chkCampo(Index).Change
End Sub

'evento - quando o check receber o foco
Private Sub chkCp_GotFocus(Index As Integer)
    chkCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    chkCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyPress(Index As Integer, KeyAscii As Integer)
    chkCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o check perder o foco
Private Sub chkCp_LostFocus(Index As Integer)
    chkCampo(Index).LostFocus
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    'Caption = LoadGasString(167870)
    vgFormID = 1634
    vgIdentTab = ""
    vgPriVez = True
    'vgTipo = TP_COMUM
    'vgCaracteristica = F_COMUM
    vgTemFiltro = False
    'vgSituacao = ACAO_NAVEGANDO
    'Set botOk.Picture = LoadPicture(LoadGasPicture(159375))
    'Set botCancela.Picture = LoadPicture(LoadGasPicture(520))
    Set btExportar = botOk
    Set lblStatus = Label(1)
    Set frameBarra = Painel(1)
    Set xRegistros = Label(2)
    Set cpDiretorioDestino = txtCampo(0)
    Set cpMes = txtCampo(1)
    Set cpAno = txtCampo(2)
    'RsOpen TbFiliais, "SELECT TOP 1 * FROM TbFiliais"
    DefineControles
    'vgTooltips.Create
    'vgTooltips.AddTool botOk, 0, LoadGasString(167875)
    'botOk.Caption = LoadGasString(167876)
    'vgTooltips.AddTool botCancela, 0, LoadGasString(167877)
    'botCancela.Caption = LoadGasString(167878)
    'Label(0).Caption = LoadGasString(167879)
    'Label(1).Caption = LoadGasString(167880)
    'Label(2).Caption = LoadGasString(167881)
    'Label(3).Caption = LoadGasString(167882)
    'Label(4).Caption = LoadGasString(167883)
    'Label(5).Caption = LoadGasString(167884)
    'Label(6).Caption = LoadGasString(167885)
    'chkCampo(0).Caption = LoadGasString(167886)
    'chkCampo(1).Caption = LoadGasString(167887)
    'chkCampo(2).Caption = LoadGasString(167888)
    'chkCampo(3).Caption = LoadGasString(167889)
    'AjustaTamanho Me
    'Executar INICIALIZACOES
    'InicializaApelidos COM_TEXTBOX
    'ExecutaPTabs
    'ExecutaPreValidacao
    FormLoad_F9
    Screen.MousePointer = vbDefault
    vgPriVez = False
End Sub


'evento - redefinido o tamanho do formulário
Private Sub Form_Resize()
    AjustaPanFundo Me
End Sub

'define as propriedades das classe dos campos do formuário
Private Sub DefineControles()

    Set txtCampo(1).CtPri = txtCp(1)
    txtCampo(1).DataType = 0
    txtCampo(1).Mask = "99"
    txtCampo(1).StuffChar = "0"

    Set txtCampo(2).CtPri = txtCp(2)
    Set txtCampo(2).CtFdo = labFdo2
    Set txtCampo(2).CtBot = bottxtCampo2
    Set bottxtCampo2.Picture = LoadPicture(LoadGasPicture(3))
    txtCampo(2).DataType = 0
    txtCampo(2).ListFields = "2013|2014|2015|2016|2017|2018|2019|2020|2021|2022|2023|2024|2025"

    Set txtCampo(3).CtPri = txtCp(3)
    Set txtCampo(3).CtFdo = labFdo3
    Set txtCampo(3).CtBot = bottxtCampo3
    Set bottxtCampo3.Picture = LoadPicture(LoadGasPicture(3))
    txtCampo(3).DataType = 1
    txtCampo(3).Mask = "99"
    txtCampo(3).BoundColumn = "IdFilial"
    txtCampo(3).ListFields = "Fantasia| CNPJ"
    Set txtCampo(3).RecordSource = New ADODB.Recordset
    RsOpen txtCampo(3).RecordSource, "SELECT TOP 1 Fantasia, CNPJ, IdFilial" + _
                    " FROM TbFiliais" + _
                    " ORDER BY Fantasia"

    Set chkCampo(0).CtPri = ChkCp(0)
    chkCampo(0).DataType = 5

    Set chkCampo(1).CtPri = ChkCp(1)
    chkCampo(1).DataType = 5

    Set chkCampo(2).CtPri = ChkCp(2)
    chkCampo(2).DataType = 5

    Set chkCampo(3).CtPri = ChkCp(3)
    chkCampo(3).DataType = 5

    Set txtCampo(0).CtPri = txtCp(0)
    Set txtCampo(0).CtFdo = labFdo0
    Set txtCampo(0).CtBot = bottxtCampo0
    Set bottxtCampo0.Picture = LoadPicture(LoadGasPicture(4))
    txtCampo(0).DataType = 0
    txtCampo(0).Mask = ""

End Sub

'executa tarefas programadas para o processo
Public Sub ProssegueExecucao()
    Dim vgNParc As Integer, x As String, vgNVez As Integer
    vgCancelou = MsgBox(LoadGasString(167871), vbExclamation + vbYesNoCancel, vgAtencao$)
    If vgCancelou = vbNo Then
        vgCancelou = True
        GoTo SaiDaSub
    ElseIf vgCancelou = vbCancel Then
        vgCancelou = True
        Exit Sub
    End If
    vgCancelou = False
    On Error GoTo DeuErro                         'vamos cercar erros...
    GeraLogAcao LoadGasString(295) + " Exportar XML"
    Screen.MousePointer = vbHourglass
    vgDb.BeginTrans                               'inicia transação
    RsOpen TbNFe, "SELECT TOP 1 * FROM TbNotaFiscalProd"
    RsOpen TbNFCe, "SELECT TOP 1 * FROM TbNFCe"
    RsOpen TbEntradaEstoque, "SELECT TOP 1 * FROM TbEntradaEstoque"
    ExecutaPTabs

    'NFe
    RsOpen TbNFe, "SELECT * FROM TbNotaFiscalProd WHERE " + FiltroInicialNFe_F9() & " ORDER BY IdNFProd"
    If TbNFe.RecordCount > 0 Then
        TbNFe.MoveFirst
        vgNVez = 0
        Do While Not TbNFe.EOF
            If ExportarXMLNFe_F9() Then
                TbEmpresa!Fantasia = ("")
            End If
            TbEmpresa.Update
            TbNFe.MoveNext
        Loop
    End If

    'NFCe
    RsOpen TbNFCe, "SELECT * FROM TbNFCe WHERE " + FiltroInicialNFCe_F9() & " ORDER BY IdNFProd"
    If TbNFCe.RecordCount > 0 Then
        TbNFCe.MoveFirst
        vgNVez = 0
        Do While Not TbNFCe.EOF
            If ExportarXMLNFCe_F9() Then
                TbEmpresa!Fantasia = ("")
            End If
            TbEmpresa.Update
            TbNFCe.MoveNext
        Loop
    End If

    'Entradas
    RsOpen TbEntradaEstoque, "SELECT * FROM TbEntradaEstoque WHERE " + FiltroInicialEntradas_F9() & " ORDER BY IDEntradaEst"
    If TbEntradaEstoque.RecordCount > 0 Then
        TbEntradaEstoque.MoveFirst
        vgNVez = 0
        Do While Not TbEntradaEstoque.EOF
            If ckEntrada Then
                If ExportarXMLEntradas_F9() Then
                    TbEmpresa!RazaoSocial = ("")
                End If
                TbEmpresa.Update
            End If
            TbEntradaEstoque.MoveNext
        Loop
    End If
    vgDb.CommitTrans                              'grava transação
    MsgBox LoadGasString(167872), vbInformation, vgAtencao$
    GoTo SaiDaSub                                 'e cai fora                                 'fora daqui

DeuErro:                                          'deu um erro
    Select Case Err                               'vamos verificar se deu algum erro
        Case 3197                                 'registro acabou de ser alterado por outro usuário
            Resume
        Case 3046, 3158, 3186, 3187, 3188, 3189, 3218, 3260
            vgNVez = vgNVez + 1                   'outro usuário acessando
            If vgNVez <= 10 Then                  'vamos tentar 10 vezes
                Delay 0.5                         'mais antes espera 1/2 segundo
                Resume                            'antes de tentar
            End If
    End Select
    If Err <> 65524 Then                          'se não for um erro de crítica
        CErr.NumErro = Err
        CErr.Show
    End If
    vgCancelou = True
    On Error Resume Next
    If TbNFe.EOF = False And TbNFe.BOF = False Then
        TbNFe.CancelUpdate
    End If
    If TbNFCe.EOF = False And TbNFCe.BOF = False Then
        TbNFCe.CancelUpdate
    End If
    If TbEntradaEstoque.EOF = False And TbEntradaEstoque.BOF = False Then
        TbEntradaEstoque.CancelUpdate
    End If
    vgDb.RollbackTrans                            'desfaz tudo

SaiDaSub:                                         'saída comum da sub...
    DoEvents
    Unload Me
    Screen.MousePointer = vbDefault               'ponteiro do mouse = default
End Sub

'evento - valida digitações e prossegue a execução
Private Sub botOk_Click()
    Dim k As String
    If vgPriVez Then Exit Sub
    vgPriVez = True
    k$ = Executar(VALIDACOES)                     'faz validaçao dos campos
    If Len(k$) > 0 Then                           'EPA!... tem campo errado
        If InStr(k$, "|") > 0 And InStr(k$, "-") > 0 And InStr(k$, "|") < InStr(k$, "-") Then
            Err.Source = Parse$(k$, "|")
            Err.Number = Val(k$)                  'deu erro
            Err.Description = Trim$(Mid$(k$, InStr(k$, "-") + 1))
            CErr.NumErro = Err.Number
            CErr.Show
        Else
            Beep                                  'manda aviso sonoro
            MsgBox k$, vbCritical, vgAtencao$     'mostra mensagem do erro
        End If
        Screen.MousePointer = vbDefault           'reseta o pointer do mouse
        vgPriVez = False
        Exit Sub                                  'e sai para continuar a ediçao ou inclusao
    End If                                        'OK. os campos estao certos
    ProssegueExecucao
    AposExportar_F9
    vgPriVez = False
End Sub

'evento - cancela formulário
Private Sub botCancela_Click()
    Unload Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    On Error Resume Next
    FinalizaForm Me
    Set btExportar = Nothing
    Set lblStatus = Nothing
    Set frameBarra = Nothing
    Set xRegistros = Nothing
    Set cpDiretorioDestino = Nothing
    Set cpMes = Nothing
    Set cpAno = Nothing
    For i = 0 To UBound(txtCampo)
        Set txtCampo(i) = Nothing
    Next
    Set chkCampo(0) = Nothing
    Set chkCampo(1) = Nothing
    Set chkCampo(2) = Nothing
    Set chkCampo(3) = Nothing
    If Not TbFiliais Is Nothing Then
        TbFiliais.Close
        Set TbFiliais = Nothing
    End If
    If Not TbNFe Is Nothing Then
        TbNFe.Close
        Set TbNFe = Nothing
    End If
    If Not TbNFCe Is Nothing Then
        TbNFCe.Close
        Set TbNFCe = Nothing
    End If
    If Not TbEntradaEstoque Is Nothing Then
        TbEntradaEstoque.Close
        Set TbEntradaEstoque = Nothing
    End If
    Set vgFG = Nothing
    Set vgTooltips = Nothing
    Set frmpExpXML = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then                'se teclou ESC
        Unload Me                                 'tira este form da memória
    End If
End Sub

'evento - quando o botão for pressionadao
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
    If vgPriVez Then Exit Sub
    Select Case Index
        Case 0
            PegaDiretorio_F9 cpDiretorioDestino
    End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click()
    txtCampo(2).botClick
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo3_Click()
    txtCampo(3).botClick
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo0_Click()
    Dim Cancel As Boolean
    txtCampo_ButtonClick 0, Cancel
    DoEvents
    If Not Cancel Then
        txtCampo(0).botClick
    End If
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtCampo(Index).MouseDown
End Sub

