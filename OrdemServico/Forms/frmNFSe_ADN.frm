VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmNFSe_ADN
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NFS-e recebidas e emitidas fora do sistema (Ambiente de Dados Nacional)"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiltro
      Caption         =   "Pesquisa"
      BeginProperty Font
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   12120
      Begin VB.ComboBox cboTipo
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1030
         Width           =   2895
      End
      Begin VB.Frame frmFiltros
         Caption         =   "Critérios"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   1260
         TabIndex        =   6
         Top             =   180
         Width           =   7335
         Begin VB.ComboBox cboNome
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Visible         =   0   'False
            Width           =   5925
         End
         Begin VB.ComboBox cboAno
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   405
            Width           =   1000
         End
         Begin VB.ComboBox cboMes
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   405
            Width           =   1500
         End
         Begin ChamaleonBtn.chameleonButton cmdPesquisar
            Height          =   315
            Left            =   6060
            TabIndex        =   16
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Pesquisar"
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
            MICON           =   "frmNFSe_ADN.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskDe
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   420
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton calDe
            Height          =   315
            Left            =   1185
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   420
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            BTYPE           =   8
            TX              =   ""
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmNFSe_ADN.frx":001C
            PICN            =   "frmNFSe_ADN.frx":0038
            PICH            =   "frmNFSe_ADN.frx":238B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox mskAte
            Height          =   315
            Left            =   1680
            TabIndex        =   13
            Top             =   420
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin ChamaleonBtn.chameleonButton calAte
            Height          =   315
            Left            =   2805
            TabIndex        =   14
            TabStop         =   0   'False
            Tag             =   "Calendario"
            Top             =   420
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            BTYPE           =   8
            TX              =   ""
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmNFSe_ADN.frx":46DE
            PICN            =   "frmNFSe_ADN.frx":46FA
            PICH            =   "frmNFSe_ADN.frx":6A4D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblAno
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano:"
            Height          =   195
            Left            =   1680
            TabIndex        =   9
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblMes
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mês:"
            Height          =   195
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   345
         End
      End
      Begin VB.Frame frmCriterios
         Caption         =   "Filtros"
         BeginProperty Font
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1155
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1155
         Begin VB.OptionButton optTodos
            Caption         =   "Todos"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   795
         End
         Begin VB.OptionButton optNome
            Caption         =   "Nome"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   795
         End
         Begin VB.OptionButton optPeriodo
            Caption         =   "Período"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   640
            Width           =   915
         End
         Begin VB.OptionButton optMensal
            Caption         =   "Mensal"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   440
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Label lblTipo
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar:"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Width           =   810
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdDocs
      Height          =   4935
      Left            =   120
      TabIndex        =   19
      Top             =   1560
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
      TabIndex        =   21
      ToolTipText     =   "Baixa do Ambiente de Dados Nacional as NFS-e e os eventos desde a última busca"
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdDanfse
      Caption         =   "&DANFSe / PDF"
      Height          =   375
      Left            =   2490
      TabIndex        =   22
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdPasta
      Caption         =   "Abrir pasta do &XML"
      Height          =   375
      Left            =   4260
      TabIndex        =   23
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdFechar
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   10545
      TabIndex        =   25
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label lblUltima
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   7980
      Width           =   12120
   End
   Begin VB.Label lblMensagem
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   120
      TabIndex        =   20
      Top             =   6555
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
Private bCarregando As Boolean
Private bNomesCarregados As Boolean
Private bComboLigado As Boolean
Private moCombo As cComboHelper        'auto preenchimento do cboNome

'Quem abre esta tela confere antes NFSe_TemADN (script 180)
'Filtros iguais aos do frmNFSe: Todos / Mensal (mes e ano) / Periodo (de-ate) / Nome (prestador ou tomador),
'pela data de emissao do documento; "Mostrar" escolhe o tipo (recebidas, emitidas fora do sistema, eventos...)
Private Sub Form_Load()
Dim i As Integer
Dim vMeses As Variant
bCarregando = True
Set moCombo = New cComboHelper
vMeses = Split("Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro", "|")
For i = 0 To 11
   cboMes.AddItem vMeses(i)
Next
cboMes.ListIndex = Month(Date) - 1
For i = Year(Date) - 5 To Year(Date) + 1
   cboAno.AddItem CStr(i)
   If i = Year(Date) Then cboAno.ListIndex = cboAno.NewIndex
Next
mskDe.Text = Format$(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yy")
mskAte.Text = Format$(Date, "dd/mm/yy")
cboTipo.AddItem "NFS-e RECEBIDAS (tomador)"
cboTipo.AddItem "Emitidas FORA do sistema (portal)"
cboTipo.AddItem "Todas as NFS-e emitidas"
cboTipo.AddItem "Todas as NFS-e"
cboTipo.AddItem "Eventos (cancelamentos etc.)"
cboTipo.ListIndex = T_RECEBIDAS
nAmbiente = Val(SQLExecutaRetorno("SELECT NFSeAmbiente FROM empresa", "NFSeAmbiente", "2"))
MontarGrid
MostrarUltima
optMensal.Value = True
AjustarCriterio
bCarregando = False
Pesquisar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set moCombo = Nothing
End Sub

'Mostra so os campos do criterio marcado. lblMes/lblAno servem aos tres criterios (mudam o texto).
Private Sub AjustarCriterio()
Dim bMensal As Boolean, bPeriodo As Boolean, bNome As Boolean
bMensal = optMensal.Value
bPeriodo = optPeriodo.Value
bNome = optNome.Value

lblMes.Visible = bMensal Or bPeriodo Or bNome
lblAno.Visible = bMensal Or bPeriodo
If bMensal Then
   lblMes.Caption = "Mês:"
   lblAno.Caption = "Ano:"
ElseIf bPeriodo Then
   lblMes.Caption = "De:"
   lblAno.Caption = "Até:"
ElseIf bNome Then
   lblMes.Caption = "Prestador ou tomador (nome ou CNPJ/CPF):"
End If
cboMes.Visible = bMensal
cboAno.Visible = bMensal
mskDe.Visible = bPeriodo
calDe.Visible = bPeriodo
mskAte.Visible = bPeriodo
calAte.Visible = bPeriodo
cboNome.Visible = bNome
cmdPesquisar.Visible = Not optTodos.Value
If bNome Then cmdPesquisar.Left = 6060 Else cmdPesquisar.Left = 3240
If bNome Then CarregarNomes
End Sub

'Prestadores e tomadores dos documentos baixados do ADN
Private Sub CarregarNomes()
Dim r As ADODB.Recordset
If bNomesCarregados Then Exit Sub
cboNome.Clear
Set r = dbData.OpenRecordset("SELECT nome FROM (SELECT EmitenteNome AS nome FROM NFSeADN WHERE (Ambiente = " & nAmbiente & ") " & _
                             "UNION SELECT TomadorNome FROM NFSeADN WHERE (Ambiente = " & nAmbiente & ")) t " & _
                             "WHERE (ISNULL(nome, '') <> '') ORDER BY nome")
Do While Not r.EOF
   cboNome.AddItem Trim$(r("nome"))
   r.MoveNext
Loop
r.Close
Set r = Nothing
If Not bComboLigado Then moCombo.AttachTo cboNome
bComboLigado = True
bNomesCarregados = True
End Sub

Private Sub optMensal_Click()
AjustarCriterio
If Not bCarregando Then cboMes.SetFocus
End Sub

Private Sub optPeriodo_Click()
AjustarCriterio
If Not bCarregando Then mskDe.SetFocus
End Sub

Private Sub optNome_Click()
AjustarCriterio
If Not bCarregando Then cboNome.SetFocus
End Sub

'Todos: sem botao Pesquisar, ja exibe tudo
Private Sub optTodos_Click()
AjustarCriterio
If Not bCarregando Then Pesquisar
End Sub

Private Function EscolherData() As Variant
Dim fCal As Calendario
Set fCal = New Calendario
fCal.Show vbModal
EscolherData = fCal.DateSelected
Unload fCal
Set fCal = Nothing
End Function

Private Sub calDe_Click()
Dim varData As Variant
varData = EscolherData()
If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub
mskDe.Text = Format$(varData, "dd/mm/yy")
End Sub

Private Sub calAte_Click()
Dim varData As Variant
varData = EscolherData()
If Not IsDate(varData) Then Exit Sub
If varData = 0 Then Exit Sub
mskAte.Text = Format$(varData, "dd/mm/yy")
End Sub

'MaskEdBox vazio fica "__/__/__"
Private Function TextoData(ByVal s As String) As String
If InStr(s, "_") > 0 Then TextoData = "" Else TextoData = s
End Function

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
   .ColWidth(C_NUM) = 1350
   .ColWidth(C_DATA) = 1100
   .ColWidth(C_EMIT) = 3300
   .ColWidth(C_TOMA) = 2350
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
Dim dIni As Date

If optMensal.Value Then
   If cboMes.ListIndex < 0 Or cboAno.ListIndex < 0 Then Exit Sub
   dIni = DateSerial(Val(cboAno.Text), cboMes.ListIndex + 1, 1)
   sDe = DataSQL(Format$(dIni, "dd/mm/yyyy"), False)
   sAte = DataSQL(Format$(DateAdd("m", 1, dIni) - 1, "dd/mm/yyyy"), True)
ElseIf optPeriodo.Value Then
   If TextoData(mskDe.Text) <> "" And Not IsDate(TextoData(mskDe.Text)) Then
      MsgBox "Data inicial inválida.", vbExclamation, "NFS-e"
      mskDe.SetFocus
      Exit Sub
   End If
   If TextoData(mskAte.Text) <> "" And Not IsDate(TextoData(mskAte.Text)) Then
      MsgBox "Data final inválida.", vbExclamation, "NFS-e"
      mskAte.SetFocus
      Exit Sub
   End If
   sDe = DataSQL(TextoData(mskDe.Text), False)
   sAte = DataSQL(TextoData(mskAte.Text), True)
End If
nTipo = cboTipo.ListIndex
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
If optNome.Value Then sNome = Trim$(cboNome.Text)
If sNome <> "" And cboNome.ListIndex >= 0 Then
   'escolhido da lista: o nome exato do prestador ou do tomador
   sSQL = sSQL & " AND ((EmitenteNome = " & NFSe_Q(sNome) & ") OR (TomadorNome = " & NFSe_Q(sNome) & "))"
ElseIf sNome <> "" Then
   'digitado: parte do nome ou o comeco do CNPJ/CPF
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
'prestadores/tomadores novos entram na lista do criterio Nome
bNomesCarregados = False
If optNome.Value Then CarregarNomes
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

'o tipo do documento filtra na hora, em qualquer criterio
Private Sub cboTipo_Click()
If Not bCarregando Then Pesquisar
End Sub

Private Sub mskDe_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: mskAte.SetFocus
End Sub

Private Sub mskAte_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub cboNome_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Pesquisar
End Sub

Private Sub mskDe_GotFocus()
SelectControl mskDe
End Sub

Private Sub mskAte_GotFocus()
SelectControl mskAte
End Sub
