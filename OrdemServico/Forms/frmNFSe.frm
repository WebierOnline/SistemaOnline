VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNFSe
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFS-e - Notas Fiscais de Serviço (padrão nacional)"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiltro
      Caption         =   "Pesquisa"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   12120
      Begin VB.TextBox txtDe
         Height          =   315
         Left            =   480
         MaxLength       =   10
         TabIndex        =   2
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox txtAte
         Height          =   315
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   4
         Top             =   285
         Width           =   1095
      End
      Begin VB.ComboBox cboSituacao
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   2055
      End
      Begin VB.TextBox txtCliente
         Height          =   315
         Left            =   6960
         TabIndex        =   8
         Top             =   285
         Width           =   3135
      End
      Begin VB.CommandButton cmdPesquisar
         Caption         =   "&Pesquisar"
         Default         =   -1  'True
         Height          =   375
         Left            =   10260
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblDe
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   375
      End
      Begin VB.Label lblAte
         Caption         =   "Até:"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   330
         Width           =   375
      End
      Begin VB.Label lblSit
         Caption         =   "Situação:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblCli
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   6330
         TabIndex        =   7
         Top             =   330
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdNotas
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdNova
      Caption         =   "&Nova NFS-e (OS)"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6780
      Width           =   1935
   End
   Begin VB.CommandButton cmdEditar
      Caption         =   "&Editar"
      Height          =   375
      Left            =   2130
      TabIndex        =   13
      Top             =   6780
      Width           =   1335
   End
   Begin VB.CommandButton cmdTransmitir
      Caption         =   "&Transmitir"
      Height          =   375
      Left            =   3540
      TabIndex        =   14
      Top             =   6780
      Width           =   1455
   End
   Begin VB.CommandButton cmdConsultar
      Caption         =   "C&onsultar"
      Height          =   375
      Left            =   5070
      TabIndex        =   15
      Top             =   6780
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir
      Caption         =   "&Imprimir / PDF"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   6780
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelarNota
      Caption         =   "Cancelar NFS-e"
      Height          =   375
      Left            =   8130
      TabIndex        =   17
      Top             =   6780
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir
      Caption         =   "E&xcluir rascunho"
      Height          =   375
      Left            =   9780
      TabIndex        =   18
      Top             =   6780
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfig
      Caption         =   "Con&figurar NFS-e"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   7260
      Width           =   1935
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   10545
      TabIndex        =   21
      Top             =   7260
      Width           =   1695
   End
   Begin VB.Label lblAmbiente
      Alignment       =   2  'Center
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   7320
      Width           =   8295
   End
   Begin VB.Label lblMensagem
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   120
      TabIndex        =   11
      Top             =   5895
      Width           =   12120
   End
End
Attribute VB_Name = "frmNFSe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Lista das NFS-e (tabela NFSe, script 178). Gerar/transmitir/consultar/cancelar/DANFSe pelo NFSeEmissor.exe (modNFSe).
'Colunas do grid: 0 Id (oculta), 1 Nº NFS-e, 2 Competência, 3 Cliente, 4 OS, 5 Valor, 6 Situação, 7 Ambiente,
'8 Mensagem (oculta), 9 código da situação (oculta), 10 PDF (oculta)
Private Const C_ID = 0, C_NUM = 1, C_COMP = 2, C_CLI = 3, C_OS = 4, C_VALOR = 5, C_SIT = 6, C_AMB = 7, C_MSG = 8, C_SITCOD = 9, C_PDF = 10

'Quem abre esta tela confere antes NFSe_BancoPronto (Unload no Form_Load de tela modal da erro 364)
Private Sub Form_Load()
txtDe.Text = Format$(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yyyy")
txtAte.Text = Format$(Date, "dd/mm/yyyy")
cboSituacao.AddItem "TODAS"
cboSituacao.AddItem "EM DIGITAÇÃO"
cboSituacao.AddItem "AUTORIZADA"
cboSituacao.AddItem "REJEITADA"
cboSituacao.AddItem "CANCELADA"
cboSituacao.AddItem "PENDENTE"
cboSituacao.ListIndex = 0
MontarGrid
MostrarAmbiente
Pesquisar
End Sub

Private Sub MontarGrid()
Dim i As Long
'MSFlexGrid exige pelo menos 1 linha alem do cabecalho: a linha 1 fica em branco quando nao ha notas
With grdNotas
   .Rows = 2
   .Cols = 11
   .FixedRows = 1
   .FixedCols = 0
   For i = 0 To .Cols - 1
      .TextMatrix(1, i) = ""
   Next
   .TextMatrix(0, C_ID) = "Id"
   .TextMatrix(0, C_NUM) = "Nº NFS-e"
   .TextMatrix(0, C_COMP) = "Competência"
   .TextMatrix(0, C_CLI) = "Cliente (tomador)"
   .TextMatrix(0, C_OS) = "OS"
   .TextMatrix(0, C_VALOR) = "Valor líquido"
   .TextMatrix(0, C_SIT) = "Situação"
   .TextMatrix(0, C_AMB) = "Ambiente"
   .ColWidth(C_ID) = 0
   .ColWidth(C_NUM) = 1100
   .ColWidth(C_COMP) = 1150
   .ColWidth(C_CLI) = 4300
   .ColWidth(C_OS) = 1900
   .ColWidth(C_VALOR) = 1300
   .ColWidth(C_SIT) = 1350
   .ColWidth(C_AMB) = 900
   .ColWidth(C_MSG) = 0
   .ColWidth(C_SITCOD) = 0
   .ColWidth(C_PDF) = 0
   .ColAlignment(C_VALOR) = flexAlignRightCenter
End With
End Sub

Private Sub MostrarAmbiente()
If Val(SQLExecutaRetorno("SELECT NFSeAmbiente FROM empresa", "NFSeAmbiente", "2")) = 1 Then
   lblAmbiente.Caption = "Ambiente da NFS-e: PRODUÇÃO (notas com valor fiscal)"
   lblAmbiente.ForeColor = &H8000&
Else
   lblAmbiente.Caption = "Ambiente da NFS-e: HOMOLOGAÇÃO (testes, sem valor fiscal) - mude em Configurar NFS-e"
   lblAmbiente.ForeColor = vbRed
End If
End Sub

Private Function DataSQL(ByVal s As String, ByVal bFim As Boolean) As String
If Not IsDate(s) Then
   DataSQL = ""
   Exit Function
End If
DataSQL = "'" & Format$(CDate(s), "yyyymmdd") & IIf(bFim, " 23:59:59", "") & "'"
End Function

Private Sub Pesquisar()
Dim r As ADODB.Recordset
Dim sSQL As String, sDe As String, sAte As String
Dim n As Long, nSit As Integer

sDe = DataSQL(txtDe.Text, False)
sAte = DataSQL(txtAte.Text, True)
sSQL = "SELECT TOP 500 IdNFSe, NumeroNFSe, DataCompetencia, TomadorNome, NumeroPedido, " & _
       "ValorServicos - DescontoIncondicionado AS Valor, Situacao, Ambiente, Mensagem, ArquivoPdf " & _
       "FROM NFSe WHERE (1 = 1)"
'Periodo pela data em que a nota foi criada (a competencia e a finalizacao da OS, pode ser de meses atras).
'Rascunho, rejeitada e pendente aparecem sempre: sao notas que ainda precisam de acao.
If sDe <> "" Or sAte <> "" Then
   sSQL = sSQL & " AND ((Situacao IN (0, 2, 4)) OR ((1 = 1)"
   If sDe <> "" Then sSQL = sSQL & " AND (DataCadastro >= " & sDe & ")"
   If sAte <> "" Then sSQL = sSQL & " AND (DataCadastro <= " & sAte & ")"
   sSQL = sSQL & "))"
End If
If cboSituacao.ListIndex > 0 Then sSQL = sSQL & " AND (Situacao = " & (cboSituacao.ListIndex - 1) & ")"
If Trim$(txtCliente.Text) <> "" Then sSQL = sSQL & " AND (TomadorNome LIKE " & NFSe_Q("%" & Trim$(txtCliente.Text) & "%") & ")"
sSQL = sSQL & " ORDER BY IdNFSe DESC"

MontarGrid
Set r = dbData.OpenRecordset(sSQL)
With grdNotas
   .Redraw = False
   Do While Not r.EOF
      n = n + 1
      If n > 1 Then .Rows = n + 1
      nSit = Val(ValidateNull(r("Situacao")))
      .TextMatrix(n, C_ID) = r("IdNFSe")
      .TextMatrix(n, C_NUM) = ValidateNull(r("NumeroNFSe"))
      .TextMatrix(n, C_COMP) = Format$(r("DataCompetencia"), "dd/mm/yyyy")
      .TextMatrix(n, C_CLI) = ValidateNull(r("TomadorNome"))
      .TextMatrix(n, C_OS) = ValidateNull(r("NumeroPedido"))
      .TextMatrix(n, C_VALOR) = Format$(ValidateNull(r("Valor")), ocMONEY)
      .TextMatrix(n, C_SIT) = NFSe_SituacaoTexto(nSit)
      .TextMatrix(n, C_AMB) = IIf(Val(ValidateNull(r("Ambiente"))) = 1, "Produção", "Teste")
      .TextMatrix(n, C_MSG) = ValidateNull(r("Mensagem"))
      .TextMatrix(n, C_SITCOD) = nSit
      .TextMatrix(n, C_PDF) = ValidateNull(r("ArquivoPdf"))
      .Row = n
      .Col = C_SIT
      Select Case nSit
         Case NFSE_AUTORIZADA: .CellForeColor = &H8000&
         Case NFSE_REJEITADA: .CellForeColor = vbRed
         Case NFSE_CANCELADA: .CellForeColor = &H808080
         Case NFSE_PENDENTE: .CellForeColor = &H80FF&
         Case Else: .CellForeColor = vbBlack
      End Select
      r.MoveNext
   Loop
   .Redraw = True
   If .Rows > 1 Then .Row = 1: .Col = 0: .ColSel = .Cols - 1
End With
r.Close
Set r = Nothing
grdNotas_RowColChange
End Sub

Private Function IdSelecionado() As Long
If grdNotas.Rows < 2 Or grdNotas.Row < 1 Then Exit Function
IdSelecionado = Val(grdNotas.TextMatrix(grdNotas.Row, C_ID))
End Function

Private Function SituacaoSelecionada() As Integer
SituacaoSelecionada = -1
If IdSelecionado() = 0 Then Exit Function
SituacaoSelecionada = Val(grdNotas.TextMatrix(grdNotas.Row, C_SITCOD))
End Function

Private Sub grdNotas_RowColChange()
Dim nSit As Integer
nSit = SituacaoSelecionada()
lblMensagem.Caption = ""
If nSit >= 0 Then lblMensagem.Caption = grdNotas.TextMatrix(grdNotas.Row, C_MSG)
cmdEditar.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA)
cmdTransmitir.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA Or nSit = NFSE_PENDENTE)
cmdConsultar.Enabled = (nSit = NFSE_PENDENTE Or nSit = NFSE_AUTORIZADA Or nSit = NFSE_REJEITADA)
cmdImprimir.Enabled = (nSit = NFSE_AUTORIZADA Or nSit = NFSE_CANCELADA)
cmdCancelarNota.Enabled = (nSit = NFSE_AUTORIZADA)
cmdExcluir.Enabled = (nSit = NFSE_DIGITACAO Or nSit = NFSE_REJEITADA)
End Sub

Private Sub grdNotas_DblClick()
Select Case SituacaoSelecionada()
   Case NFSE_DIGITACAO, NFSE_REJEITADA: cmdEditar_Click
   Case NFSE_AUTORIZADA, NFSE_CANCELADA: cmdImprimir_Click
End Select
End Sub

Private Sub cmdPesquisar_Click()
Pesquisar
End Sub

Private Sub cmdNova_Click()
Dim lId As Long
frmNFSe_OS.IdGerado = 0
frmNFSe_OS.Show vbModal
lId = frmNFSe_OS.IdGerado
Unload frmNFSe_OS
If lId = 0 Then Exit Sub
Pesquisar
SelecionarId lId
If frmNFSe_Editar.Editar(lId) Then TransmitirId lId
Pesquisar
SelecionarId lId
End Sub

Private Sub SelecionarId(ByVal lId As Long)
Dim i As Long
For i = 1 To grdNotas.Rows - 1
   If Val(grdNotas.TextMatrix(i, C_ID)) = lId Then
      grdNotas.Row = i
      grdNotas.Col = 0
      grdNotas.ColSel = grdNotas.Cols - 1
      If Not grdNotas.RowIsVisible(i) Then grdNotas.TopRow = i
      grdNotas_RowColChange
      Exit Sub
   End If
Next
End Sub

Private Sub cmdEditar_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If frmNFSe_Editar.Editar(lId) Then TransmitirId lId
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdTransmitir_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If ShowMsg("Transmitir a NFS-e " & grdNotas.TextMatrix(grdNotas.Row, C_OS) & " de " & grdNotas.TextMatrix(grdNotas.Row, C_CLI) & _
           " (R$ " & grdNotas.TextMatrix(grdNotas.Row, C_VALOR) & ")?", vbQuestion + vbYesNo + vbDefaultButton1) <> vbYes Then Exit Sub
TransmitirId lId
Pesquisar
SelecionarId lId
End Sub

'Transmite e mostra o resultado (autorizada: oferece abrir a DANFSe)
Private Sub TransmitirId(ByVal lId As Long)
Dim oRes As Object
Dim sStatus As String
sStatus = NFSe_Executar("acao=transmitir" & vbCrLf & "id=" & lId, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível transmitir:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
   Exit Sub
End If
Select Case Val(NFSe_Valor(oRes, "situacao"))
   Case NFSE_AUTORIZADA
      If MsgBox("NFS-e AUTORIZADA!" & vbCrLf & vbCrLf & "Número: " & NFSe_Valor(oRes, "numero") & vbCrLf & _
                "Chave: " & NFSe_Valor(oRes, "chave") & vbCrLf & vbCrLf & "Abrir a DANFSe (PDF) agora?", vbInformation + vbYesNo, "NFS-e") = vbYes Then
         NFSe_AbrirArquivo NFSe_Valor(oRes, "pdf")
      End If
   Case NFSE_REJEITADA
      MsgBox "A NFS-e foi REJEITADA pelo Sefin:" & vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem") & vbCrLf & vbCrLf & _
             "Corrija em Editar (ou na Configuração / cadastro do cliente) e transmita de novo.", vbExclamation, "NFS-e"
   Case NFSE_PENDENTE
      MsgBox "Sem resposta do Sefin: a NFS-e ficou PENDENTE." & vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem") & vbCrLf & vbCrLf & _
             "Use Consultar daqui a alguns minutos (não emita outra nota para as mesmas OS).", vbExclamation, "NFS-e"
   Case Else
      MsgBox NFSe_Valor(oRes, "mensagem"), vbInformation, "NFS-e"
End Select
End Sub

Private Sub cmdConsultar_Click()
Dim oRes As Object
Dim sStatus As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
sStatus = NFSe_Executar("acao=consultar" & vbCrLf & "id=" & lId, oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível consultar:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
Else
   MsgBox "Situação: " & NFSe_SituacaoTexto(Val(NFSe_Valor(oRes, "situacao"))) & _
          IIf(NFSe_Valor(oRes, "numero") <> "", vbCrLf & "Número: " & NFSe_Valor(oRes, "numero"), "") & _
          IIf(NFSe_Valor(oRes, "mensagem") <> "", vbCrLf & vbCrLf & NFSe_Valor(oRes, "mensagem"), ""), vbInformation, "NFS-e"
End If
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdImprimir_Click()
Dim oRes As Object
Dim sStatus As String, sPdf As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
sPdf = grdNotas.TextMatrix(grdNotas.Row, C_PDF)
If sPdf <> "" Then
   If Dir$(sPdf) = "" Then sPdf = ""
End If
If sPdf = "" Then
   'PDF apagado ou de outro computador: gera de novo a partir do XML gravado no banco
   sStatus = NFSe_Executar("acao=pdf" & vbCrLf & "id=" & lId, oRes)
   If Left$(sStatus, 2) <> "OK" Then
      MsgBox "Não foi possível gerar a DANFSe:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
      Exit Sub
   End If
   sPdf = NFSe_Valor(oRes, "pdf")
End If
NFSe_AbrirArquivo sPdf
End Sub

Private Sub cmdCancelarNota_Click()
Dim oRes As Object
Dim sStatus As String
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
frmNFSe_Cancelar.Confirmado = False
frmNFSe_Cancelar.lblNota.Caption = "NFS-e nº " & grdNotas.TextMatrix(grdNotas.Row, C_NUM) & " - " & grdNotas.TextMatrix(grdNotas.Row, C_CLI) & _
                                   " - R$ " & grdNotas.TextMatrix(grdNotas.Row, C_VALOR)
frmNFSe_Cancelar.Show vbModal
If Not frmNFSe_Cancelar.Confirmado Then
   Unload frmNFSe_Cancelar
   Exit Sub
End If
sStatus = NFSe_Executar("acao=cancelar" & vbCrLf & "id=" & lId & vbCrLf & "motivo=" & frmNFSe_Cancelar.Motivo & vbCrLf & _
                        "justificativa=" & Replace(Replace(frmNFSe_Cancelar.Justificativa, vbCr, " "), vbLf, " "), oRes)
Unload frmNFSe_Cancelar
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível cancelar:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
Else
   MsgBox "NFS-e cancelada." & vbCrLf & vbCrLf & "As OS dessa nota ficam livres para uma nova NFS-e.", vbInformation, "NFS-e"
End If
Pesquisar
SelecionarId lId
End Sub

Private Sub cmdExcluir_Click()
Dim lId As Long
lId = IdSelecionado()
If lId = 0 Then Exit Sub
If ShowMsg("Excluir o rascunho da NFS-e " & grdNotas.TextMatrix(grdNotas.Row, C_OS) & "?" & vbCrLf & vbCrLf & _
           "As OS voltam a ficar disponíveis para gerar outra NFS-e.", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
'so rascunho ou rejeitada (o Sefin nao gerou nota): NFSeItens sai junto (ON DELETE CASCADE)
dbData.Execute "DELETE FROM NFSe WHERE (IdNFSe = " & lId & ") AND (Situacao IN (0, 2));"
Pesquisar
End Sub

Private Sub cmdConfig_Click()
frmNFSe_Config.Show vbModal
MostrarAmbiente
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub txtDe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub txtAte_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub
