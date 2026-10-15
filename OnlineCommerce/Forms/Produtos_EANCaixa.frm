VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "chamaleonbtn.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Produtos_EANCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EAN Caixa / Embalagem"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCarregando 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5340
      Top             =   3200
   End
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame fraFiltro 
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   10800
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
         Height          =   975
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   6075
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "Filtrar"
            Default         =   -1  'True
            Height          =   315
            Left            =   4860
            TabIndex        =   18
            Top             =   480
            Width           =   1050
         End
         Begin VB.TextBox txtFiltro 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   4680
         End
         Begin VB.OptionButton optPalavra 
            Caption         =   "Palavra"
            Height          =   255
            Left            =   2340
            TabIndex        =   14
            Top             =   210
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.OptionButton optPalavraDupla 
            Caption         =   "Palavra dupla"
            Height          =   255
            Left            =   3360
            TabIndex        =   13
            Top             =   210
            Visible         =   0   'False
            Width           =   1300
         End
         Begin VB.Label lblFiltro 
            AutoSize        =   -1  'True
            Caption         =   "EAN:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
      End
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
         Height          =   975
         Left            =   60
         TabIndex        =   8
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   180
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optEAN 
            Caption         =   "EAN"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            Top             =   450
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optDescricao 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   660
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkSoSemCaixa 
         Caption         =   "Mostrar so quem nao tem EAN de caixa cadastrado"
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         Top             =   1080
         Width           =   4455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4950
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   8731
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame fraCarregando 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3060
      TabIndex        =   3
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
         TabIndex        =   4
         Top             =   90
         Width           =   4680
      End
   End
   Begin ChamaleonBtn.chameleonButton cmdVerCadastro 
      Height          =   315
      Left            =   7620
      TabIndex        =   17
      ToolTipText     =   "Adiciona"
      Top             =   6420
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Exibir Cadastro"
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
      MICON           =   "Produtos_EANCaixa.frx":0000
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
      Height          =   315
      Left            =   9300
      TabIndex        =   19
      ToolTipText     =   "Adiciona"
      Top             =   6420
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Fechar"
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
      MICON           =   "Produtos_EANCaixa.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblLegenda 
      AutoSize        =   -1  'True
      Caption         =   "Clique numa celula azul pra editar (EAN / UND / fracao da caixa). "
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   6420
      Width           =   3960
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   6780
      Width           =   120
   End
End
Attribute VB_Name = "Produtos_EANCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Manutencao em massa dos campos Produtos.EANCaixa (EAN da CAIXA fechada - campo se
'chamava EANEmbalagem antes), Produtos.UnidMedCX (unidade da caixa, ex CX/FD/DZ),
'Produtos.Fracionamento (quantos itens vem numa caixa) e Produtos.EAN (EAN da
'EMBALAGEM/unidade - o que o cliente leva na prateleira). Grid editavel: clica na
'celula, digita, Enter ou clica fora pra salvar (grava direto, linha a linha).
'
'Colunas do grid: 0=CODIGO(oculta) 1=DESCRICAO 2=EAN CAIXA(editavel) 3=UND(editavel,
'Produtos.UnidMedCX) 4=FRACAO(editavel, qtd por caixa) - grupo CAIXA, fundo azul claro
'| 5=EAN EMBALAGEM 6=UND(Produtos.UNID_MEDIDA) 7=FRACAO(fixo "1") - grupo EMBALAGEM,
'fundo verde claro, somente leitura. So as colunas 2, 3 e 4 sao editaveis.

Private iDots As Integer
Private iRow As Integer
Private iCol As Integer
Private bCancelarEdicao As Boolean

Private Sub Form_Load()
    iDots = 0
    fraCarregando.Visible = False
    ConfigurarGrid
    optEAN.Value = True
    optPalavraDupla.Value = True
    AtualizarTipoConsulta
End Sub

'SetFocus nao pode ser chamado dentro do Form_Load (form ainda nao esta visivel/ativo -
'da erro 5). Form_Activate dispara logo depois, com o form ja na tela.
Private Sub Form_Activate()
    txtFiltro.SetFocus
End Sub

'Controla visibilidade/legendas do bloco de filtro conforme optTodos/optEAN/optDescricao
Private Sub AtualizarTipoConsulta()
    Dim bMostraFiltro As Boolean
    bMostraFiltro = Not optTodos.Value

    lblFiltro.Visible = bMostraFiltro
    txtFiltro.Visible = bMostraFiltro
    optPalavra.Visible = bMostraFiltro
    optPalavraDupla.Visible = bMostraFiltro
    cmdFiltrar.Visible = bMostraFiltro
    chkSoSemCaixa.Visible = optTodos.Value

    If optEAN.Value Then
        lblFiltro.Caption = "EAN:"
        optPalavra.Caption = "Caixa"
        optPalavraDupla.Caption = "Embalagem"
    ElseIf optDescricao.Value Then
        lblFiltro.Caption = "Descricao:"
        optPalavra.Caption = "Palavra"
        optPalavraDupla.Caption = "Palavra dupla"
    End If
End Sub

'Unico lugar que carrega os 500 registros sem filtro nenhum - clique explicito em Todos
Private Sub optTodos_Click()
    AtualizarTipoConsulta
    CarregarGrid
End Sub

Private Sub optEAN_Click()
    AtualizarTipoConsulta
    optPalavraDupla.Value = True
    txtFiltro.Text = ""
    Grid.rows = 1
    lblTotal.Caption = ""
    txtFiltro.SetFocus
End Sub

Private Sub optDescricao_Click()
    AtualizarTipoConsulta
    optPalavra.Value = True
    txtFiltro.Text = ""
    Grid.rows = 1
    lblTotal.Caption = ""
    txtFiltro.SetFocus
End Sub

Private Sub ConfigurarGrid()
    With Grid
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = 4200    'Descricao
        .ColWidth(2) = 1500    'EAN caixa
        .ColWidth(3) = 600     'Und caixa (fixo CX)
        .ColWidth(4) = 900     'Fracao caixa
        .ColWidth(5) = 1500    'EAN embalagem
        .ColWidth(6) = 600     'Und embalagem
        .ColWidth(7) = 900     'Fracao embalagem (fixo 1)
        .Row = 0: .Col = 1: .Text = "DESCRICAO"
        .Row = 0: .Col = 2: .Text = "EAN CAIXA"
        .Row = 0: .Col = 3: .Text = "UND"
        .Row = 0: .Col = 4: .Text = "FRACAO"
        .Row = 0: .Col = 5: .Text = "EAN EMBALAGEM"
        .Row = 0: .Col = 6: .Text = "UND"
        .Row = 0: .Col = 7: .Text = "FRACAO"
        'tinge o cabecalho pra marcar os 2 grupos (caixa x embalagem)
        Dim jc As Integer
        For jc = 2 To 4
            .Row = 0: .Col = jc: .CellBackColor = &HFFE8CC   'azul claro - CAIXA
        Next jc
        For jc = 5 To 7
            .Row = 0: .Col = jc: .CellBackColor = &HE0FFE0   'verde claro - EMBALAGEM
        Next jc
        .AllowUserResizing = 1
        .FixedCols = 0
        .rows = 1
    End With
End Sub

Private Sub CarregarGrid()
    'EAN/Descricao exigem um termo digitado - evita carregar tudo sem querer.
    'Os 500 registros sem filtro so aparecem com optTodos.
    If Not optTodos.Value And Trim(txtFiltro.Text) = "" Then
        Grid.rows = 1
        lblTotal.Caption = ""
        Exit Sub
    End If

    fraCarregando.Visible = True
    fraCarregando.ZOrder 0
    tmrCarregando.Enabled = True
    Me.Refresh
    DoEvents

    Grid.rows = 1

    Dim w As String
    w = "WHERE ATIVO = 1"

    If optEAN.Value Then
        If Trim(txtFiltro.Text) <> "" Then
            Dim sEAN As String
            sEAN = Replace(Trim(txtFiltro.Text), "'", "''")
            If optPalavra.Value Then
                'Caixa
                w = w & " AND EANCaixa LIKE '%" & sEAN & "%'"
            Else
                'Embalagem
                w = w & " AND EAN LIKE '%" & sEAN & "%'"
            End If
        End If
    ElseIf optDescricao.Value Then
        If Trim(txtFiltro.Text) <> "" Then
            If optPalavra.Value Then
                w = w & " AND DESCRICAO LIKE '%" & Replace(Trim(txtFiltro.Text), "'", "''") & "%'"
            Else
                'Palavra dupla: separa no primeiro espaco - as 2 palavras tem que aparecer,
                'cada uma em qualquer parte do nome (nao precisa estar juntas/na ordem digitada)
                Dim sTextoDupla As String, sPal1 As String, sPal2 As String, iEspaco As Integer
                sTextoDupla = Trim(txtFiltro.Text)
                iEspaco = InStr(sTextoDupla, " ")
                If iEspaco > 0 Then
                    sPal1 = Trim(Left(sTextoDupla, iEspaco - 1))
                    sPal2 = Trim(Mid(sTextoDupla, iEspaco + 1))
                Else
                    sPal1 = sTextoDupla
                    sPal2 = sTextoDupla
                End If
                w = w & " AND DESCRICAO LIKE '%" & Replace(sPal1, "'", "''") & "%'" & _
                        " AND DESCRICAO LIKE '%" & Replace(sPal2, "'", "''") & "%'"
            End If
        End If
    End If
    'optTodos: nao acrescenta filtro de texto nenhum - mostra tudo (ate o limite de 500)

    If chkSoSemCaixa.Value = 1 Then
        w = w & " AND (EANCaixa IS NULL OR EANCaixa = '' OR EANCaixa = 'SEM GTIN')"
    End If

    Dim sql As String
    sql = "SELECT TOP 500 CODIGO, DESCRICAO, ISNULL(EANCaixa,'') AS EANCaixa, " & _
          "ISNULL(UnidMedCX,'CX') AS UndCaixa, ISNULL(Fracionamento,1) AS FracaoCaixa, " & _
          "ISNULL(EAN,'') AS EANEmb, ISNULL(UNID_MEDIDA,'') AS Unidade FROM Produtos " & w & " ORDER BY DESCRICAO"

    Dim rP As ADODB.Recordset
    RsOpen rP, sql

    Dim r As Integer
    r = 1
    Do While Not rP.EOF
        With Grid
            .rows = r + 1
            .Row = r: .Col = 0: .Text = CStr(rP("CODIGO"))
            .Row = r: .Col = 1: .Text = Trim(rP("DESCRICAO") & "")
            .Row = r: .Col = 2: .Text = Trim(rP("EANCaixa") & "")
            .Row = r: .Col = 3: .Text = Trim(rP("UndCaixa") & "")
            .Row = r: .Col = 4: .Text = FormatNumber(rP("FracaoCaixa"), 0)
            .Row = r: .Col = 5: .Text = Trim(rP("EANEmb") & "")
            .Row = r: .Col = 6: .Text = Trim(rP("Unidade") & "")
            .Row = r: .Col = 7: .Text = "1"
            Dim jc As Integer
            For jc = 2 To 4
                .Row = r: .Col = jc: .CellBackColor = &HFFF6EA
            Next jc
            For jc = 5 To 7
                .Row = r: .Col = jc: .CellBackColor = &HEFFFEF
            Next jc
            .Row = r: .Col = 5: .CellForeColor = &H808080
            .Row = r: .Col = 6: .CellForeColor = &H808080
            .Row = r: .Col = 7: .CellForeColor = &H808080
        End With
        r = r + 1
        rP.MoveNext
    Loop
    If rP.State <> 0 Then rP.Close

    lblTotal.Caption = (r - 1) & " produto(s) encontrado(s)" & IIf(r - 1 = 500, " (limitado a 500 - refine o filtro)", "")

    tmrCarregando.Enabled = False
    fraCarregando.Visible = False
End Sub

Private Sub cmdFiltrar_Click()
    CarregarGrid
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        CarregarGrid
    End If
End Sub

Private Sub chkSoSemCaixa_Click()
    CarregarGrid
End Sub

'Abre o txtEdit flutuante sobre a celula clicada, se for uma coluna editavel
Private Sub Grid_Click()
    Dim bEditavel As Boolean
    bEditavel = False
    Select Case Grid.Col
        Case 2, 3, 4
            bEditavel = True
    End Select
    If Not bEditavel Or Grid.Row < 1 Then Exit Sub

    iRow = Grid.Row
    iCol = Grid.Col
    bCancelarEdicao = False
    txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight
    txtEdit.Text = Grid.TextMatrix(iRow, iCol)
    txtEdit.Visible = True
    txtEdit.ZOrder 0
    txtEdit.SetFocus
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then          'Enter confirma
        KeyAscii = 0
        Grid.SetFocus
    ElseIf KeyAscii = 27 Then      'Esc cancela
        KeyAscii = 0
        bCancelarEdicao = True
        Grid.SetFocus
    End If
End Sub

'Valida e grava a celula editada direto em Produtos. Cada coluna editavel grava
'num campo diferente: 2=EANCaixa 3=UnidMedCX 4=Fracionamento
Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
    If bCancelarEdicao Then
        bCancelarEdicao = False
        Exit Sub
    End If

    Dim sVal As String
    sVal = Trim(txtEdit.Text)
    Dim lCodigo As Long
    lCodigo = Val(Grid.TextMatrix(iRow, 0))
    If lCodigo <= 0 Then Exit Sub

    Dim msgErro As String

    Select Case iCol

        Case 2 'EAN CAIXA
            sVal = Replace(sVal, " ", "")
            If sVal = "" Then sVal = "SEM GTIN"
            If sVal <> "SEM GTIN" Then
                If Len(sVal) = 8 Or Len(sVal) = 13 Then
                    Dim i As Integer, soma As Long, pesoD As Integer
                    soma = 0
                    For i = 1 To Len(sVal) - 1
                        If Len(sVal) = 13 Then
                            If i Mod 2 = 1 Then pesoD = 1 Else pesoD = 3
                        Else
                            If i Mod 2 = 1 Then pesoD = 3 Else pesoD = 1
                        End If
                        soma = soma + Val(Mid(sVal, i, 1)) * pesoD
                    Next i
                    If (10 - (soma Mod 10)) Mod 10 <> Val(Right(sVal, 1)) Then
                        MsgBox "EAN invalido: digito verificador incorreto.", vbExclamation
                        Exit Sub
                    End If
                End If
            End If
            msgErro = SQLExecuta("UPDATE Produtos SET EANCaixa = '" & Replace(sVal, "'", "''") & "' WHERE CODIGO = " & lCodigo)
            If Not Vazio(msgErro) Then
                MsgBox "Erro ao salvar: " & msgErro, vbCritical
                Exit Sub
            End If
            Grid.TextMatrix(iRow, iCol) = sVal

        Case 3 'UND (caixa) - mesma lista de unidades de volume grande do cboUnidMedCX
            sVal = UCase(sVal)
            If sVal = "" Then sVal = "CX"
            Select Case sVal
                Case "BJ", "CT", "CX", "DI", "DZ", "EX", "FD", "GL", "PA", "PT", "SC"
                    'ok
                Case Else
                    MsgBox "Unidade invalida. Use uma das unidades de caixa/fardo: " & _
                           "BJ, CT, CX, DI, DZ, EX, FD, GL, PA, PT, SC", vbExclamation
                    Exit Sub
            End Select
            msgErro = SQLExecuta("UPDATE Produtos SET UnidMedCX = '" & sVal & "' WHERE CODIGO = " & lCodigo)
            If Not Vazio(msgErro) Then
                MsgBox "Erro ao salvar: " & msgErro, vbCritical
                Exit Sub
            End If
            Grid.TextMatrix(iRow, iCol) = sVal

        Case 4 'FRACAO (caixa)
            Dim dFrac As Double
            dFrac = Val(Replace(sVal, ",", "."))
            If dFrac <= 0 Then dFrac = 1
            msgErro = SQLExecuta("UPDATE Produtos SET Fracionamento = " & FSQL(dFrac) & " WHERE CODIGO = " & lCodigo)
            If Not Vazio(msgErro) Then
                MsgBox "Erro ao salvar: " & msgErro, vbCritical
                Exit Sub
            End If
            Grid.TextMatrix(iRow, iCol) = FormatNumber(dFrac, 0)

    End Select
End Sub

'Abre o cadastro do produto selecionado na linha do grid (coluna 0 = CODIGO)
Private Sub cmdVerCadastro_Click()
    Dim idxVer As Integer
    idxVer = Grid.Row
    If idxVer < 1 Then
        MsgBox "Selecione um produto na lista.", vbExclamation
        Exit Sub
    End If

    Dim lCodVer As Long
    lCodVer = Val(Grid.TextMatrix(idxVer, 0))
    If lCodVer <= 0 Then Exit Sub

    Load Produtos_Cadastro
    Produtos_Cadastro.SSTab1.Tab = 0
    Produtos_Cadastro.EditarProduto lCodVer
    Produtos_Cadastro.Show vbModal, Me
    Unload Produtos_Cadastro

    CarregarGrid
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
