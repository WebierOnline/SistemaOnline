VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNFSe_ADN
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFS-e recebidas e emitidas fora do sistema (Ambiente de Dados Nacional)"
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
      Begin VB.ComboBox cboTipo
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   2895
      End
      Begin VB.TextBox txtNome
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         ToolTipText     =   "Parte do nome ou do CNPJ/CPF do prestador ou do tomador"
         Top             =   285
         Width           =   2535
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
      Begin VB.Label lblTipo
         Caption         =   "Mostrar:"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   330
         Width           =   615
      End
      Begin VB.Label lblNome
         Caption         =   "Nome:"
         Height          =   255
         Left            =   6960
         TabIndex        =   7
         Top             =   330
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdDocs
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
   Begin VB.CommandButton cmdBuscar
      Caption         =   "&Buscar novas no ADN"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Baixa do Ambiente de Dados Nacional as NFS-e e os eventos desde a última busca"
      Top             =   6780
      Width           =   2295
   End
   Begin VB.CommandButton cmdDanfse
      Caption         =   "&DANFSe / PDF"
      Height          =   375
      Left            =   2490
      TabIndex        =   13
      Top             =   6780
      Width           =   1695
   End
   Begin VB.CommandButton cmdPasta
      Caption         =   "Abrir pasta do &XML"
      Height          =   375
      Left            =   4260
      TabIndex        =   14
      Top             =   6780
      Width           =   1935
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   10545
      TabIndex        =   16
      Top             =   6780
      Width           =   1695
   End
   Begin VB.Label lblUltima
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   12120
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
Attribute VB_Name = "frmNFSe_ADN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Documentos da distribuicao do ADN (tabela NFSeADN, script 180), baixados pelo NFSeEmissor (acao=distribuir):
'NFS-e RECEBIDAS (empresa como tomador/intermediario), emitidas (inclusive as feitas direto no portal) e eventos.
'Colunas: 0 NSU (oculta), 1 Nº, 2 Emissão, 3 Prestador, 4 Tomador, 5 Valor, 6 Situação, 7 Origem,
'8 descricao/chave (oculta, vai no rodape), 9 tipo do documento (oculta), 10 situacao (oculta)
Private Const C_NSU = 0, C_NUM = 1, C_DATA = 2, C_EMIT = 3, C_TOMA = 4, C_VALOR = 5, C_SIT = 6, C_ORIGEM = 7, C_DESC = 8, C_TIPO = 9, C_SITCOD = 10

Private Const T_RECEBIDAS = 0, T_FORA = 1, T_EMITIDAS = 2, T_TODAS = 3, T_EVENTOS = 4

Private nAmbiente As Integer

'Quem abre esta tela confere antes NFSe_TemADN (script 180)
Private Sub Form_Load()
txtDe.Text = Format$(DateSerial(Year(Date), Month(Date) - 2, 1), "dd/mm/yyyy")
txtAte.Text = Format$(Date, "dd/mm/yyyy")
cboTipo.AddItem "NFS-e RECEBIDAS (tomador)"
cboTipo.AddItem "Emitidas FORA do sistema (portal)"
cboTipo.AddItem "Todas as NFS-e emitidas"
cboTipo.AddItem "Todas as NFS-e"
cboTipo.AddItem "Eventos (cancelamentos etc.)"
cboTipo.ListIndex = T_RECEBIDAS
nAmbiente = Val(SQLExecutaRetorno("SELECT NFSeAmbiente FROM empresa", "NFSeAmbiente", "2"))
MontarGrid
MostrarUltima
Pesquisar
End Sub

Private Sub MontarGrid()
Dim i As Long
With grdDocs
   .Rows = 2
   .Cols = 11
   .FixedRows = 1
   .FixedCols = 0
   For i = 0 To .Cols - 1
      .TextMatrix(1, i) = ""
   Next
   .TextMatrix(0, C_NSU) = "NSU"
   .TextMatrix(0, C_NUM) = "Nº NFS-e"
   .TextMatrix(0, C_DATA) = "Emissão"
   .TextMatrix(0, C_EMIT) = "Prestador (emitente)"
   .TextMatrix(0, C_TOMA) = "Tomador"
   .TextMatrix(0, C_VALOR) = "Valor"
   .TextMatrix(0, C_SIT) = "Situação"
   .TextMatrix(0, C_ORIGEM) = "Origem"
   .ColWidth(C_NSU) = 0
   .ColWidth(C_NUM) = 1000
   .ColWidth(C_DATA) = 1100
   .ColWidth(C_EMIT) = 3300
   .ColWidth(C_TOMA) = 2700
   .ColWidth(C_VALOR) = 1150
   .ColWidth(C_SIT) = 1100
   .ColWidth(C_ORIGEM) = 1700
   .ColWidth(C_DESC) = 0
   .ColWidth(C_TIPO) = 0
   .ColWidth(C_SITCOD) = 0
   .ColAlignment(C_VALOR) = flexAlignRightCenter
End With
End Sub

Private Sub MostrarUltima()
Dim r As ADODB.Recordset
Set r = dbData.OpenRecordset("SELECT CONVERT(varchar(20), MAX(NSU)) AS nsu, MAX(DataRecebimento) AS dh, COUNT(*) AS qtd FROM NFSeADN WHERE (Ambiente = " & nAmbiente & ")")
If Val(ValidateNull(r("qtd"))) = 0 Then
   lblUltima.Caption = "Nenhum documento baixado ainda" & IIf(nAmbiente = 1, "", " (ambiente de HOMOLOGAÇÃO)") & _
                       ": clique em Buscar novas no ADN (a primeira busca traz todo o histórico do CNPJ)."
Else
   lblUltima.Caption = "Último documento recebido em " & Format$(r("dh"), "dd/mm/yyyy hh:nn") & " (NSU " & r("nsu") & ", " & _
                       r("qtd") & " documento(s) no total)" & IIf(nAmbiente = 1, "", " - ambiente de HOMOLOGAÇÃO")
End If
r.Close
Set r = Nothing
End Sub

Private Function DataSQL(ByVal s As String, ByVal bFim As Boolean) As String
If Not IsDate(s) Then
   DataSQL = ""
   Exit Function
End If
DataSQL = "'" & Format$(CDate(s), "yyyymmdd") & IIf(bFim, " 23:59:59", "") & "'"
End Function

Private Function NomeEvento(ByVal sTipo As String) As String
Select Case sTipo
   Case "101101": NomeEvento = "Cancelamento"
   Case "105102": NomeEvento = "Cancelamento por substituição"
   Case "101103": NomeEvento = "Pedido de análise fiscal (cancelamento)"
   Case "105104": NomeEvento = "Cancelamento DEFERIDO pela prefeitura"
   Case "105105": NomeEvento = "Cancelamento INDEFERIDO pela prefeitura"
   Case "": NomeEvento = "Evento"
   Case Else: NomeEvento = "Evento " & sTipo
End Select
End Function

Private Sub Pesquisar()
Dim r As ADODB.Recordset
Dim sSQL As String, sDe As String, sAte As String, sNome As String, sDoc As String
Dim n As Long, nTipo As Integer, nPapel As Integer, nSit As Integer
Dim bEvento As Boolean

nTipo = cboTipo.ListIndex
sDe = DataSQL(txtDe.Text, False)
sAte = DataSQL(txtAte.Text, True)
sSQL = "SELECT TOP 1000 CONVERT(varchar(20), NSU) AS NSUTxt, ChaveAcesso, TipoDocumento, TipoEvento, Papel, Situacao, NumeroNFSe, DataEmissao, EmitenteDoc, EmitenteNome, " & _
       "TomadorDoc, TomadorNome, ValorServico, CodTribNac, Descricao, Municipio, IdNFSe FROM NFSeADN WHERE (Ambiente = " & nAmbiente & ")"
Select Case nTipo
   Case T_RECEBIDAS: sSQL = sSQL & " AND (TipoDocumento = 'NFSE') AND (Papel IN (2, 3))"
   Case T_FORA: sSQL = sSQL & " AND (TipoDocumento = 'NFSE') AND (Papel = 1) AND (IdNFSe IS NULL)"
   Case T_EMITIDAS: sSQL = sSQL & " AND (TipoDocumento = 'NFSE') AND (Papel = 1)"
   Case T_TODAS: sSQL = sSQL & " AND (TipoDocumento = 'NFSE')"
   Case T_EVENTOS: sSQL = sSQL & " AND (TipoDocumento <> 'NFSE')"
End Select
If sDe <> "" Then sSQL = sSQL & " AND (DataEmissao >= " & sDe & ")"
If sAte <> "" Then sSQL = sSQL & " AND (DataEmissao <= " & sAte & ")"
sNome = Trim$(txtNome.Text)
If sNome <> "" Then
   sDoc = SoNumeros(sNome)
   sSQL = sSQL & " AND ((EmitenteNome LIKE " & NFSe_Q("%" & sNome & "%") & ") OR (TomadorNome LIKE " & NFSe_Q("%" & sNome & "%") & ")"
   If Len(sDoc) >= 4 Then sSQL = sSQL & " OR (EmitenteDoc LIKE " & NFSe_Q(sDoc & "%") & ") OR (TomadorDoc LIKE " & NFSe_Q(sDoc & "%") & ")"
   sSQL = sSQL & ")"
End If
sSQL = sSQL & " ORDER BY DataEmissao DESC, NSU DESC"

MontarGrid
Set r = dbData.OpenRecordset(sSQL)
With grdDocs
   .Redraw = False
   Do While Not r.EOF
      n = n + 1
      If n > 1 Then .Rows = n + 1
      bEvento = (ValidateNull(r("TipoDocumento")) <> "NFSE")
      nPapel = Val(ValidateNull(r("Papel")))
      nSit = Val(ValidateNull(r("Situacao")))
      .TextMatrix(n, C_NSU) = r("NSUTxt")
      .TextMatrix(n, C_NUM) = ValidateNull(r("NumeroNFSe"))
      If Not IsNull(r("DataEmissao")) Then .TextMatrix(n, C_DATA) = Format$(r("DataEmissao"), "dd/mm/yyyy")
      .TextMatrix(n, C_EMIT) = ValidateNull(r("EmitenteNome"))
      If .TextMatrix(n, C_EMIT) = "" Then .TextMatrix(n, C_EMIT) = ValidateNull(r("EmitenteDoc"))
      .TextMatrix(n, C_TOMA) = ValidateNull(r("TomadorNome"))
      If .TextMatrix(n, C_TOMA) = "" Then .TextMatrix(n, C_TOMA) = ValidateNull(r("TomadorDoc"))
      If Not IsNull(r("ValorServico")) Then .TextMatrix(n, C_VALOR) = Format$(r("ValorServico"), ocMONEY)
      .TextMatrix(n, C_TIPO) = IIf(bEvento, "E", "N")
      .TextMatrix(n, C_SITCOD) = nSit
      If bEvento Then
         .TextMatrix(n, C_SIT) = ""
         .TextMatrix(n, C_ORIGEM) = NomeEvento(ValidateNull(r("TipoEvento")))
         .TextMatrix(n, C_DESC) = NomeEvento(ValidateNull(r("TipoEvento"))) & " da NFS-e de chave " & ValidateNull(r("ChaveAcesso"))
      Else
         .TextMatrix(n, C_SIT) = IIf(nSit = 3, "CANCELADA", "NORMAL")
         Select Case nPapel
            Case 1: .TextMatrix(n, C_ORIGEM) = IIf(IsNull(r("IdNFSe")), "Emitida FORA do sistema", "Emitida pelo sistema")
            Case 2: .TextMatrix(n, C_ORIGEM) = "Recebida (tomador)"
            Case 3: .TextMatrix(n, C_ORIGEM) = "Recebida (intermediário)"
            Case Else: .TextMatrix(n, C_ORIGEM) = "Outro"
         End Select
         .TextMatrix(n, C_DESC) = IIf(ValidateNull(r("CodTribNac")) <> "", "Serviço " & r("CodTribNac") & ": ", "") & ValidateNull(r("Descricao")) & _
                                  IIf(ValidateNull(r("Municipio")) <> "", " | Local: " & r("Municipio"), "") & " | Chave: " & ValidateNull(r("ChaveAcesso"))
         .Row = n
         .Col = C_SIT
         .CellForeColor = IIf(nSit = 3, &H808080, &H8000&)
         .Col = C_ORIGEM
         If nPapel = 1 And IsNull(r("IdNFSe")) Then .CellForeColor = &H80FF&
      End If
      r.MoveNext
   Loop
   .Redraw = True
   If .Rows > 1 Then .Row = 1: .Col = 0: .ColSel = .Cols - 1
End With
r.Close
Set r = Nothing
grdDocs_RowColChange
End Sub

Private Function SoNumeros(ByVal s As String) As String
Dim i As Integer
For i = 1 To Len(s)
   If Mid$(s, i, 1) Like "#" Then SoNumeros = SoNumeros & Mid$(s, i, 1)
Next
End Function

Private Function NsuSelecionado() As String
If grdDocs.Rows < 2 Or grdDocs.Row < 1 Then Exit Function
NsuSelecionado = grdDocs.TextMatrix(grdDocs.Row, C_NSU)
End Function

Private Function NFSeSelecionada() As Boolean
If NsuSelecionado() = "" Then Exit Function
NFSeSelecionada = (grdDocs.TextMatrix(grdDocs.Row, C_TIPO) = "N")
End Function

Private Sub grdDocs_RowColChange()
lblMensagem.Caption = ""
If NsuSelecionado() <> "" Then lblMensagem.Caption = grdDocs.TextMatrix(grdDocs.Row, C_DESC)
cmdDanfse.Enabled = NFSeSelecionada()
cmdPasta.Enabled = NFSeSelecionada()
End Sub

Private Sub grdDocs_DblClick()
If NFSeSelecionada() Then cmdDanfse_Click
End Sub

Private Sub cmdPesquisar_Click()
Pesquisar
End Sub

Private Sub cmdBuscar_Click()
Dim oRes As Object
Dim sStatus As String
sStatus = NFSe_Executar("acao=distribuir", oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível buscar as NFS-e no Ambiente de Dados Nacional:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
Else
   MsgBox NFSe_Valor(oRes, "mensagem") & _
          IIf(Val(NFSe_Valor(oRes, "atualizadas")) > 0, vbCrLf & vbCrLf & "Confira a tela de NFS-e: notas do sistema mudaram de situação (cancelamento/análise fiscal).", ""), _
          vbInformation, "NFS-e"
   If Val(NFSe_Valor(oRes, "recebidas")) > 0 And cboTipo.ListIndex <> T_RECEBIDAS Then cboTipo.ListIndex = T_RECEBIDAS
End If
MostrarUltima
Pesquisar
End Sub

'Gera XML + DANFSe em <DiretorioXML>\NFSe\ADN\AAAAMM\ e devolve o caminho do PDF ("" se falhou)
Private Function GerarArquivos() As String
Dim oRes As Object
Dim sStatus As String
sStatus = NFSe_Executar("acao=pdfadn" & vbCrLf & "nsu=" & NsuSelecionado(), oRes)
If Left$(sStatus, 2) <> "OK" Then
   MsgBox "Não foi possível gerar a DANFSe:" & vbCrLf & vbCrLf & Mid$(sStatus, 7), vbExclamation, "NFS-e"
   Exit Function
End If
GerarArquivos = NFSe_Valor(oRes, "pdf")
End Function

Private Sub cmdDanfse_Click()
Dim sPdf As String
If Not NFSeSelecionada() Then Exit Sub
sPdf = GerarArquivos()
If sPdf <> "" Then NFSe_AbrirArquivo sPdf
End Sub

Private Sub cmdPasta_Click()
Dim sPdf As String
Dim oShell As Object
If Not NFSeSelecionada() Then Exit Sub
sPdf = GerarArquivos()
If sPdf = "" Then Exit Sub
On Error Resume Next
Set oShell = CreateObject("WScript.Shell")
oShell.Run "explorer.exe /select," & Chr(34) & Replace(sPdf, "-danfse.pdf", "-nfse.xml") & Chr(34), 1, False
Set oShell = Nothing
End Sub

Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub cboTipo_Click()
If Me.Visible Then Pesquisar
End Sub

Private Sub txtDe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub txtAte_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub
