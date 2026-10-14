VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form VinculoXMLProduto_Consulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vinculos de Produtos (XML)"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCarregando 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   7740
      Top             =   5340
   End
   Begin VB.Frame fraFiltro 
      Height          =   1230
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   13920
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   2040
         TabIndex        =   19
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optFornecedor 
            Caption         =   "Fornecedor"
            Height          =   255
            Left            =   180
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optDescricao 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   180
            TabIndex        =   21
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton optEAN 
            Caption         =   "EAN"
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   180
            TabIndex        =   25
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Critérios"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   4020
         TabIndex        =   12
         Top             =   120
         Width           =   4875
         Begin VB.ComboBox cboFiltroFornecedor 
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   4020
         End
         Begin VB.TextBox txtFiltroDesc 
            Height          =   315
            Left            =   195
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   4560
         End
         Begin VB.TextBox txtFiltroEAN 
            Height          =   315
            Left            =   180
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.OptionButton optPalavraDupla 
            Caption         =   "Palavra dupla"
            Height          =   255
            Left            =   3360
            TabIndex        =   14
            Top             =   210
            Visible         =   0   'False
            Width           =   1300
         End
         Begin VB.OptionButton optPalavra 
            Caption         =   "Palavra"
            Height          =   255
            Left            =   2460
            TabIndex        =   13
            Top             =   210
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.Label lblFiltroFornecedor 
            AutoSize        =   -1  'True
            Caption         =   "Fornecedor:"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblFiltroDesc 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblFiltroEAN 
            AutoSize        =   -1  'True
            Caption         =   "EAN:"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Consulta de Produtos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1875
         Begin VB.OptionButton optCompra 
            Caption         =   "Compra"
            Height          =   255
            Left            =   180
            TabIndex        =   11
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optInterno 
            Caption         =   "Interno"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            Top             =   480
            Value           =   -1  'True
            Width           =   800
         End
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Filtrar"
         Default         =   -1  'True
         Height          =   825
         Left            =   8940
         TabIndex        =   0
         Top             =   210
         Width           =   1260
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5640
      Left            =   60
      TabIndex        =   1
      Top             =   1260
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   9948
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdDesvincular 
      Caption         =   "Desvincular"
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   6960
      Width           =   2400
   End
   Begin VB.CommandButton cmdAlterarProduto 
      Caption         =   "Alterar Produto Interno"
      Height          =   450
      Left            =   2580
      TabIndex        =   3
      Top             =   6960
      Width           =   2700
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   450
      Left            =   11580
      TabIndex        =   4
      Top             =   6960
      Width           =   2400
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   4350
      TabIndex        =   6
      Top             =   3200
      Visible         =   0   'False
      Width           =   4800
      Begin VB.Label lblCarregando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   7
         Top             =   90
         Width           =   4680
      End
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5400
      TabIndex        =   5
      Top             =   7110
      Width           =   120
   End
End
Attribute VB_Name = "VinculoXMLProduto_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tela de consulta/manutencao da tabela VinculoXMLProduto - mostra TODOS os vinculos ja
'gravados (qualquer fornecedor, qualquer nota), diferente do "Mostrar vinculos" de dentro
'do Entrada_Estoque (que so olha pra nota atual). Util pra achar/corrigir vinculo errado
'sem precisar reabrir a nota onde ele foi feito.
'
'Colunas do grid: 0=IDFornecedor(oculta) 1=Fornecedor 2=Cod.Forn 3=EAN.Forn 4=Desc.Forn
'5=Und.Forn (lado esquerdo, dados da COMPRA/fornecedor - fundo azul claro) | 6=Cod.Interno
'7=EAN.Interno 8=Desc.Interno 9=Und.Interno (lado direito, PRODUTO INTERNO - fundo verde claro)
'
'Filtros: optCompra/optInterno decide se EAN e Descricao buscam nos dados do FORNECEDOR
'(v.EANEmbalagem/v.xProd) ou do PRODUTO INTERNO (p.EAN/p.DESCRICAO) - padrao Interno.
'optPalavra/optPalavraDupla decide se a Descricao busca 1 termo ou 2 termos (cada um em
'qualquer parte do nome, nao precisa ser sequencia exata) - padrao Palavra.
'Fornecedor (combo) nao depende do toggle Compra/Interno - sempre filtra f.razao.
Private iDots As Integer

'Preenchidas por quem abre esta tela de fora (ex: cmdBuscarVinculo do
'frmVinculoProdutoXML) pra abrir ja filtrado - deixar em branco pra abrir sem filtro
'inicial nenhum (padrao = Fornecedor, conforme o design da tela)
Public sEANInicial As String
Public sFornecedorInicial As String

Private Sub Form_Load()
    iDots = 0
    fraCarregando.Visible = False
    ConfigurarGrid
    CarregarFornecedores

    'NAO le sEANInicial/sFornecedorInicial aqui - o VB6 dispara Form_Load no
    'momento do "Load NomeDoForm", ANTES de quem chamou ter chance de setar essas
    'propriedades. Quem abre com filtro inicial deve chamar AplicarFiltroInicial
    'manualmente, depois de setar as propriedades e antes do .Show.
    AtualizarTipoConsulta
    CarregarGrid
End Sub

'Chamar DEPOIS de setar sEANInicial/sFornecedorInicial e ANTES do .Show - aplica o
'filtro inicial e recarrega o grid. Precisa ser assim (nao dentro do Form_Load)
'porque o Load do VB6 dispara o Form_Load antes do caller setar as propriedades.
Public Sub AplicarFiltroInicial()
    If Not Vazio(sEANInicial) Then
        optEAN.Value = True
        optCompra.Value = True
        txtFiltroEAN.Text = sEANInicial
        AtualizarTipoConsulta
        CarregarGrid
    ElseIf Not Vazio(sFornecedorInicial) Then
        optFornecedor.Value = True
        cboFiltroFornecedor.Text = sFornecedorInicial
        AtualizarTipoConsulta
        CarregarGrid
    End If
End Sub

'Mostra so os campos do tipo de consulta selecionado (EAN/Descricao/Fornecedor) e
'esconde os outros dois - optEAN/optDescricao/optFornecedor sao mutuamente exclusivos
Private Sub AtualizarTipoConsulta()
    lblFiltroEAN.Visible = optEAN.Value
    txtFiltroEAN.Visible = optEAN.Value

    lblFiltroDesc.Visible = optDescricao.Value
    txtFiltroDesc.Visible = optDescricao.Value
    optPalavra.Visible = optDescricao.Value
    optPalavraDupla.Visible = optDescricao.Value

    lblFiltroFornecedor.Visible = optFornecedor.Value
    cboFiltroFornecedor.Visible = optFornecedor.Value
End Sub

Private Sub optEAN_Click()
    AtualizarTipoConsulta
End Sub

Private Sub optDescricao_Click()
    AtualizarTipoConsulta
End Sub

Private Sub optFornecedor_Click()
    AtualizarTipoConsulta
End Sub

Private Sub optTodos_Click()
    AtualizarTipoConsulta
End Sub

Private Sub CarregarFornecedores()
    cboFiltroFornecedor.Clear
    cboFiltroFornecedor.AddItem ""
    Dim rF As ADODB.Recordset
    RsOpen rF, "SELECT DISTINCT f.razao FROM VinculoXMLProduto v " & _
                "INNER JOIN fornecedor f ON f.codigo = v.IDFornecedor " & _
                "ORDER BY f.razao"
    Do While Not rF.EOF
        cboFiltroFornecedor.AddItem Trim(rF("razao") & "")
        rF.MoveNext
    Loop
    If rF.State <> 0 Then rF.Close
End Sub

Private Sub ConfigurarGrid()
    With Grid
        .Cols = 10
        .ColWidth(0) = 0
        .ColWidth(1) = 2000    'Fornecedor
        .ColWidth(2) = 715     'Cod (fornecedor) - 50% do original, +10% depois
        .ColWidth(3) = 1400    'EAN (fornecedor)
        .ColWidth(4) = 3432    'Descricao (fornecedor) - +30% do original, +10% depois
        .ColWidth(5) = 560     'Und (fornecedor) - -20% do original (700)
        .ColWidth(6) = 550     'Cod (interno) - 50% do original, +10% depois
        .ColWidth(7) = 1400    'EAN (interno)
        .ColWidth(8) = 3432    'Descricao (interno) - +30% do original, +10% depois
        .ColWidth(9) = 560     'Und (interno) - -20% do original (700)
        .Row = 0: .Col = 1: .Text = "FORNECEDOR"
        .Row = 0: .Col = 2: .Text = "COD"
        .Row = 0: .Col = 3: .Text = "EAN"
        .Row = 0: .Col = 4: .Text = "DESCRICAO"
        .Row = 0: .Col = 5: .Text = "UND"
        .Row = 0: .Col = 6: .Text = "COD"
        .Row = 0: .Col = 7: .Text = "EAN"
        .Row = 0: .Col = 8: .Text = "DESCRICAO"
        .Row = 0: .Col = 9: .Text = "UND"
        'tinge o cabecalho pra marcar os 2 grupos (compra x produto interno)
        Dim jc As Integer
        For jc = 1 To 5
            .Row = 0: .Col = jc: .CellBackColor = &HFFE8CC   'azul claro - dados do FORNECEDOR/compra
        Next jc
        For jc = 6 To 9
            .Row = 0: .Col = jc: .CellBackColor = &HE0FFE0   'verde claro - PRODUTO INTERNO
        Next jc
        .AllowUserResizing = 1
        .SelectionMode = 1
        .rows = 1
    End With
End Sub

Private Sub CarregarGrid()
    Dim rV As ADODB.Recordset
    Dim sql As String, w As String
    Dim r As Integer

    fraCarregando.Visible = True
    fraCarregando.ZOrder 0
    tmrCarregando.Enabled = True
    Me.Refresh
    DoEvents

    Grid.rows = 1

    w = "WHERE v.IDProduto <> 0"

    'So o tipo de consulta ATIVO (EAN/Descricao/Fornecedor) entra na query - evita que
    'texto esquecido num campo escondido filtre por engano
    If optEAN.Value Then
        'EAN - busca no lado COMPRA (XML do fornecedor) ou INTERNO (cadastro), conforme o toggle
        If Trim(txtFiltroEAN.Text) <> "" Then
            Dim sEANF As String
            sEANF = Replace(Trim(txtFiltroEAN.Text), "'", "''")
            If optCompra.Value Then
                w = w & " AND v.EANEmbalagem LIKE '%" & sEANF & "%'"
            Else
                w = w & " AND p.EAN LIKE '%" & sEANF & "%'"
            End If
        End If

    ElseIf optDescricao.Value Then
        'Descricao - mesmo toggle Compra/Interno decide o campo; Palavra/Palavra dupla decide
        'se busca 1 termo ou 2 termos (cada um em qualquer posicao do nome)
        If Trim(txtFiltroDesc.Text) <> "" Then
            Dim sCampoDesc As String
            sCampoDesc = IIf(optCompra.Value, "v.xProd", "p.DESCRICAO")

            If optPalavra.Value Then
                w = w & " AND " & sCampoDesc & " LIKE '%" & Replace(Trim(txtFiltroDesc.Text), "'", "''") & "%'"
            Else
                'Palavra dupla: separa no primeiro espaco - as 2 palavras tem que aparecer,
                'cada uma em qualquer parte do nome (nao precisa estar juntas/na ordem digitada)
                Dim sTextoDupla As String, sPal1 As String, sPal2 As String, iEspaco As Integer
                sTextoDupla = Trim(txtFiltroDesc.Text)
                iEspaco = InStr(sTextoDupla, " ")
                If iEspaco > 0 Then
                    sPal1 = Trim(Left(sTextoDupla, iEspaco - 1))
                    sPal2 = Trim(Mid(sTextoDupla, iEspaco + 1))
                Else
                    sPal1 = sTextoDupla
                    sPal2 = sTextoDupla
                End If
                w = w & " AND " & sCampoDesc & " LIKE '%" & Replace(sPal1, "'", "''") & "%'" & _
                        " AND " & sCampoDesc & " LIKE '%" & Replace(sPal2, "'", "''") & "%'"
            End If
        End If

    ElseIf optFornecedor.Value Then
        'Fornecedor - sempre pelo nome do fornecedor, independente do toggle Compra/Interno
        If Trim(cboFiltroFornecedor.Text) <> "" Then
            w = w & " AND f.razao LIKE '%" & Replace(Trim(cboFiltroFornecedor.Text), "'", "''") & "%'"
        End If
    End If

    sql = "SELECT TOP 500 v.IDFornecedor, ISNULL(f.razao,'(fornecedor nao encontrado)') AS Fornecedor, " & _
          "v.cProd, ISNULL(v.EANEmbalagem,'') AS EANForn, ISNULL(v.xProd,'') AS DescForn, ISNULL(v.uCom,'') AS UndForn, " & _
          "v.IDProduto, ISNULL(p.EAN,'') AS EANInt, ISNULL(p.DESCRICAO,'') AS DescInt, ISNULL(p.UNID_MEDIDA,'') AS UndInt " & _
          "FROM VinculoXMLProduto v " & _
          "INNER JOIN Produtos p ON p.Codigo = v.IDProduto " & _
          "LEFT JOIN fornecedor f ON f.codigo = v.IDFornecedor " & _
          w & " ORDER BY Fornecedor, DescForn"

    RsOpen rV, sql
    r = 1
    Do While Not rV.EOF
        With Grid
            .rows = r + 1
            .Row = r: .Col = 0: .Text = CStr(rV("IDFornecedor"))
            .Row = r: .Col = 1: .Text = Trim(rV("Fornecedor") & "")
            .Row = r: .Col = 2: .Text = Trim(rV("cProd") & "")
            .Row = r: .Col = 3: .Text = Trim(rV("EANForn") & "")
            .Row = r: .Col = 4: .Text = Trim(rV("DescForn") & "")
            .Row = r: .Col = 5: .Text = Trim(rV("UndForn") & "")
            .Row = r: .Col = 6: .Text = CStr(rV("IDProduto"))
            .Row = r: .Col = 7: .Text = Trim(rV("EANInt") & "")
            .Row = r: .Col = 8: .Text = Trim(rV("DescInt") & "")
            .Row = r: .Col = 9: .Text = Trim(rV("UndInt") & "")
            Dim jc2 As Integer
            For jc2 = 1 To 5
                .Row = r: .Col = jc2: .CellBackColor = &HFFF6EA
            Next jc2
            For jc2 = 6 To 9
                .Row = r: .Col = jc2: .CellBackColor = &HEFFFEF
            Next jc2
        End With
        r = r + 1
        rV.MoveNext
    Loop
    If rV.State <> 0 Then rV.Close

    lblTotal.Caption = (r - 1) & " vinculo(s) encontrado(s)" & IIf(r - 1 = 500, " (limitado a 500 - refine o filtro)", "")

    tmrCarregando.Enabled = False
    fraCarregando.Visible = False
End Sub

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

Private Sub txtFiltroEAN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CarregarGrid
End Sub

Private Sub txtFiltroDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then CarregarGrid
End Sub

Private Sub cmdDesvincular_Click()
    If Grid.Row < 1 Then
        MsgBox "Selecione um vinculo na lista.", vbExclamation
        Exit Sub
    End If

    Dim lIDForn As Long, sCProd As String, sDescForn As String, sDescInt As String
    lIDForn = Val(Grid.TextMatrix(Grid.Row, 0))
    sCProd = Grid.TextMatrix(Grid.Row, 2)
    sDescForn = Grid.TextMatrix(Grid.Row, 4)
    sDescInt = Grid.TextMatrix(Grid.Row, 8)

    If MsgBox("Desfazer este vinculo?" & vbCrLf & vbCrLf & _
              "Fornecedor/Item: " & sCProd & " - " & sDescForn & vbCrLf & _
              "Produto interno : " & sDescInt & vbCrLf & vbCrLf & _
              "Isso NAO altera nenhuma nota ou estoque ja lancado - so impede que esse " & _
              "codigo do fornecedor continue reconhecendo esse produto automaticamente " & _
              "nas proximas entradas.", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Desvinculo") = vbNo Then Exit Sub

    Dim msgErro As String
    msgErro = SQLExecuta("UPDATE VinculoXMLProduto SET IDProduto = 0, EANProduto = '', " & _
                          "Fracionamento = 1, CustoUnitario = 0, DataAtualizacao = GETDATE() " & _
                          "WHERE IDFornecedor = " & lIDForn & " AND cProd = '" & Replace(sCProd, "'", "''") & "'")
    If Not Vazio(msgErro) Then
        MsgBox "Erro ao desvincular: " & msgErro, vbCritical
        Exit Sub
    End If

    CarregarGrid
End Sub

Private Sub cmdAlterarProduto_Click()
    If Grid.Row < 1 Then
        MsgBox "Selecione um vinculo na lista.", vbExclamation
        Exit Sub
    End If

    Dim lCodProd As Long
    lCodProd = Val(Grid.TextMatrix(Grid.Row, 6))
    If lCodProd <= 0 Then Exit Sub

    Load Produtos_Cadastro
    Produtos_Cadastro.SSTab1.Tab = 0
    Produtos_Cadastro.EditarProduto lCodProd
    Produtos_Cadastro.Show vbModal, Me
    Unload Produtos_Cadastro

    CarregarGrid
End Sub

Private Sub Grid_DblClick()
    cmdAlterarProduto_Click
End Sub

Private Sub cmdFechar_Click()
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub tmrCarregando_Timer()
    iDots = (iDots Mod 3) + 1
    lblCarregando.Caption = "Carregando" & String(iDots, ".")
End Sub
