path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Adicionar cboEdit ao layout do form ────────────────────────────────────
content = sub('Add cboEdit layout',
    '   Begin VB.TextBox txtEdit \r\n',
    '   Begin VB.ComboBox cboEdit \r\n'
    '      Height          =   315\r\n'
    '      Left            =   2520\r\n'
    '      TabIndex        =   50\r\n'
    '      Top             =   2640\r\n'
    '      Visible         =   0   \'False\r\n'
    '      Width           =   1500\r\n'
    '   End\r\n'
    '   Begin VB.TextBox txtEdit \r\n',
    content)

# ── 2: Declarar tipoEmpresa no modulo ─────────────────────────────────────────
content = sub('Declarar tipoEmpresa',
    'Private iRow As Long, iCol As Long\r\n',
    'Private iRow As Long, iCol As Long\r\n'
    'Private tipoEmpresa As Long\r\n',
    content)

# ── 3: Carregar tipoEmpresa no Form_Load ──────────────────────────────────────
content = sub('Form_Load tipoEmpresa',
    'Set moCombo = New cComboHelper\r\n'
    '\r\n'
    '\'tipo de venda = 1 simples e 2 multiplus pre\xe7os\r\n'
    'Set cCfg = sysConfig("TIPOVALORVENDA")',

    'Set moCombo = New cComboHelper\r\n'
    'tipoEmpresa = CLng(sysConfig("TIPO_EMPRESA").Value)\r\n'
    '\r\n'
    '\'tipo de venda = 1 simples e 2 multiplus pre\xe7os\r\n'
    'Set cCfg = sysConfig("TIPOVALORVENDA")',
    content)

# ── 4: Adicionar var_tags nas queries (produtos.) ─────────────────────────────
content = sub('queries add var_tags produtos',
    'produtos.categoria AS var_cat, produtos.fabricante',
    'produtos.categoria AS var_cat, produtos.TAGS AS var_tags, produtos.fabricante',
    content, replace_all=True)

# ── 5: Adicionar var_tags na query com alias p. ───────────────────────────────
content = sub('query add var_tags p.',
    'p.categoria AS var_cat, p.fabricante',
    'p.categoria AS var_cat, p.TAGS AS var_tags, p.fabricante',
    content)

# ── 6: Formatar_Grid .Cols 11 -> 12 ──────────────────────────────────────────
content = sub('Formatar_Grid Cols 11->12',
    '      .Cols = 11\r\n      .rows = 2',
    '      .Cols = 12\r\n      .rows = 2',
    content)

# ── 7: Formatar_Grid ColWidths 7-10 ──────────────────────────────────────────
content = sub('Formatar_Grid ColWidths add TAG',
    '      .ColWidth(4) = 5200\r\n'
    '      .ColWidth(5) = 1600\r\n'
    '      .ColWidth(6) = 800\r\n'
    '      .ColWidth(7) = 1750\r\n'
    '      .ColWidth(8) = 800\r\n'
    '      .ColWidth(9) = 1000\r\n'
    '      .ColWidth(10) = 1000',

    '      .ColWidth(4) = 5200\r\n'
    '      .ColWidth(5) = 1600\r\n'
    '      .ColWidth(6) = 800\r\n'
    '      .ColWidth(7) = 1750\r\n'
    '      .ColWidth(8) = 1200\r\n'
    '      .ColWidth(9) = 800\r\n'
    '      .ColWidth(10) = 1000\r\n'
    '      .ColWidth(11) = 1000',
    content)

# ── 8: Formatar_Grid cabecalhos col 7-10 ─────────────────────────────────────
content = sub('Formatar_Grid headers add TAG',
    '      .TextMatrix(0, 7) = "CATEGORIA"\r\n'
    '      .TextMatrix(0, 8) = "LOCAL"\r\n'
    '      .TextMatrix(0, 9) = "ESTOQUE"\r\n'
    '      .TextMatrix(0, 10) = "VENDA"',

    '      .TextMatrix(0, 7) = "CATEGORIA"\r\n'
    '      .TextMatrix(0, 8) = "TAG"\r\n'
    '      .TextMatrix(0, 9) = "LOCAL"\r\n'
    '      .TextMatrix(0, 10) = "ESTOQUE"\r\n'
    '      .TextMatrix(0, 11) = "VENDA"',
    content)

# ── 9: Formatar_Grid dados col 7-10 ──────────────────────────────────────────
content = sub('Formatar_Grid data add TAG',
    '            .TextMatrix(.rows - 1, 7) = Format$(ValidateNull(rTabela("var_cat")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_Local"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_quant"))\r\n'
    '            .TextMatrix(.rows - 1, 10) = Format$(ValidateNull(rTabela("venda")), ocMONEY)',

    '            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_cat"))\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_tags"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_Local"))\r\n'
    '            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_quant"))\r\n'
    '            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)',
    content)

# ── 10: Formatar_Grid_Fiscal .Cols 14 -> 15 ───────────────────────────────────
content = sub('Formatar_Grid_Fiscal Cols 14->15',
    '      .Cols = 14\r\n      .rows = 2',
    '      .Cols = 15\r\n      .rows = 2',
    content)

# ── 11: Formatar_Grid_Fiscal ColWidths 7-13 ───────────────────────────────────
content = sub('Formatar_Grid_Fiscal ColWidths add TAG',
    '      .ColWidth(4) = 4200\r\n'
    '      .ColWidth(5) = 1600\r\n'
    '      .ColWidth(6) = 800\r\n'
    '      .ColWidth(7) = 1750\r\n'
    '      .ColWidth(8) = 800\r\n'
    '      .ColWidth(9) = 1000\r\n'
    '      .ColWidth(10) = 1000\r\n'
    '      .ColWidth(11) = 1000\r\n'
    '      .ColWidth(12) = 1100\r\n'
    '      .ColWidth(13) = 1100',

    '      .ColWidth(4) = 4200\r\n'
    '      .ColWidth(5) = 1600\r\n'
    '      .ColWidth(6) = 800\r\n'
    '      .ColWidth(7) = 1750\r\n'
    '      .ColWidth(8) = 1200\r\n'
    '      .ColWidth(9) = 800\r\n'
    '      .ColWidth(10) = 1000\r\n'
    '      .ColWidth(11) = 1000\r\n'
    '      .ColWidth(12) = 1000\r\n'
    '      .ColWidth(13) = 1100\r\n'
    '      .ColWidth(14) = 1100',
    content)

# ── 12: Formatar_Grid_Fiscal cabecalhos col 7-13 ──────────────────────────────
content = sub('Formatar_Grid_Fiscal headers add TAG',
    '      .TextMatrix(0, 7) = "CATEGORIA"\r\n'
    '      .TextMatrix(0, 8) = "LOCAL"\r\n'
    '      .TextMatrix(0, 9) = "FISCAL"\r\n'
    '      .TextMatrix(0, 10) = "ESTOQUE"\r\n'
    '      .TextMatrix(0, 11) = "VENDA"\r\n'
    '      .TextMatrix(0, 12) = "CUSTO"\r\n'
    '      .TextMatrix(0, 13) = "T.FISCAL',

    '      .TextMatrix(0, 7) = "CATEGORIA"\r\n'
    '      .TextMatrix(0, 8) = "TAG"\r\n'
    '      .TextMatrix(0, 9) = "LOCAL"\r\n'
    '      .TextMatrix(0, 10) = "FISCAL"\r\n'
    '      .TextMatrix(0, 11) = "ESTOQUE"\r\n'
    '      .TextMatrix(0, 12) = "VENDA"\r\n'
    '      .TextMatrix(0, 13) = "CUSTO"\r\n'
    '      .TextMatrix(0, 14) = "T.FISCAL',
    content)

# ── 13: Formatar_Grid_Fiscal dados col 7-12 + VarTotalGrid ────────────────────
content = sub('Formatar_Grid_Fiscal data add TAG',
    '            .TextMatrix(.rows - 1, 7) = Format$(ValidateNull(rTabela("var_cat")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_Local"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_EstoqueFiscal"))\r\n'
    '            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_quant"))\r\n'
    '            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("venda")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("custo")), ocMONEY)\r\n'
    '            \r\n'
    '            VarTotalGrid = .TextMatrix(.rows - 1, 12) * .TextMatrix(.rows - 1, 9)\r\n'
    '            .TextMatrix(.rows - 1, 13) = Format(VarTotalGrid, ocMONEY)',

    '            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_cat"))\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_tags"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_Local"))\r\n'
    '            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_EstoqueFiscal"))\r\n'
    '            .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela("var_quant"))\r\n'
    '            .TextMatrix(.rows - 1, 12) = Format$(ValidateNull(rTabela("venda")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 13) = Format$(ValidateNull(rTabela("custo")), ocMONEY)\r\n'
    '            \r\n'
    '            VarTotalGrid = .TextMatrix(.rows - 1, 13) * .TextMatrix(.rows - 1, 10)\r\n'
    '            .TextMatrix(.rows - 1, 14) = Format(VarTotalGrid, ocMONEY)',
    content)

# ── 14: Grid_Click — reescrever com cboEdit para cols 7 e 8 ──────────────────
content = sub('Grid_Click rewrite',
    'Private Sub Grid_Click()\r\n'
    '\'Criado por mim\r\n'
    'Dim i As Integer\r\n'
    'Dim ColLimite As Integer\r\n'
    '\r\n'
    'If optMostrarFiscal.Value = True Then\r\n'
    '    ColLimite = 9\r\n'
    'Else\r\n'
    '    ColLimite = 8\r\n'
    'End If\r\n'
    '\r\n'
    'For i = 3 To ColLimite\r\n'
    '   If Grid.ColSel = i Then\r\n'
    '      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '      txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)\r\n'
    '      txtEdit.Visible = True\r\n'
    '      txtEdit.SetFocus\r\n'
    '      txtEdit.SelStart = 0\r\n'
    '      txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    '      iRow = Grid.Row\r\n'
    '      iCol = Grid.Col\r\n'
    '   End If\r\n'
    'Next\r\n'
    'End Sub',

    'Private Sub Grid_Click()\r\n'
    'Dim ColLimite As Integer\r\n'
    'Dim rCbo As ADODB.Recordset\r\n'
    '\r\n'
    'If optMostrarFiscal.Value = True Then\r\n'
    '    ColLimite = 10\r\n'
    'Else\r\n'
    '    ColLimite = 9\r\n'
    'End If\r\n'
    '\r\n'
    'iRow = Grid.Row\r\n'
    'iCol = Grid.Col\r\n'
    '\r\n'
    'If iCol < 3 Or iCol > ColLimite Then Exit Sub\r\n'
    '\r\n'
    'Select Case iCol\r\n'
    '   Case 7\r\n'
    '      cboEdit.Clear\r\n'
    '      Set rCbo = dbData.OpenRecordset("SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;")\r\n'
    '      Do While Not rCbo.EOF\r\n'
    '         cboEdit.AddItem ValidateNull(rCbo("Categoria"))\r\n'
    '         rCbo.MoveNext\r\n'
    '      Loop\r\n'
    '      If rCbo.State <> 0 Then rCbo.Close\r\n'
    '      Set rCbo = Nothing\r\n'
    '      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '      cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '      cboEdit.Visible = True\r\n'
    '      cboEdit.SetFocus\r\n'
    '   Case 8\r\n'
    '      If Len(Trim(Grid.TextMatrix(iRow, 7))) = 0 Then\r\n'
    '         MsgBox "Defina a categoria do produto antes de editar a tag.", vbExclamation, "Tag sem categoria"\r\n'
    '         Exit Sub\r\n'
    '      End If\r\n'
    '      cboEdit.Clear\r\n'
    '      Set rCbo = dbData.OpenRecordset("SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;")\r\n'
    '      Do While Not rCbo.EOF\r\n'
    '         cboEdit.AddItem ValidateNull(rCbo("Tags"))\r\n'
    '         rCbo.MoveNext\r\n'
    '      Loop\r\n'
    '      If rCbo.State <> 0 Then rCbo.Close\r\n'
    '      Set rCbo = Nothing\r\n'
    '      cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '      cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '      cboEdit.Visible = True\r\n'
    '      cboEdit.SetFocus\r\n'
    '   Case Else\r\n'
    '      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '      txtEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '      txtEdit.Visible = True\r\n'
    '      txtEdit.SetFocus\r\n'
    '      txtEdit.SelStart = 0\r\n'
    '      txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    'End Select\r\n'
    'End Sub',
    content)

# ── 15: Adicionar handlers cboEdit apos txtEdit_LostFocus ────────────────────
content = sub('Add cboEdit handlers',
    'Private Sub txtEdit_LostFocus()\r\n'
    '\'criado por mim\r\n'
    'If iCol = 6 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n'
    '\r\n'
    'Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\r\n'
    'txtEdit.Visible = False\r\n'
    'End Sub',

    'Private Sub txtEdit_LostFocus()\r\n'
    '\'criado por mim\r\n'
    'If iCol = 6 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n'
    '\r\n'
    'Grid.TextMatrix(iRow, iCol) = IIf(txtEdit.Text = "", 0, txtEdit.Text)\r\n'
    'txtEdit.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdit_LostFocus()\r\n'
    '   If iRow > 0 Then\r\n'
    '      Dim sVal As String\r\n'
    '      sVal = Trim(cboEdit.Text)\r\n'
    '      If iCol = 8 Then sVal = UCase(sVal)\r\n'
    '      Grid.TextMatrix(iRow, iCol) = sVal\r\n'
    '   End If\r\n'
    '   cboEdit.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdit_KeyPress(KeyAscii As Integer)\r\n'
    '   If iCol = 8 Then\r\n'
    '      If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32\r\n'
    '   End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdit_Change()\r\n'
    '   If iCol = 7 Or iCol = 8 Then\r\n'
    '      Dim pos As Integer\r\n'
    '      pos = cboEdit.SelStart\r\n'
    '      cboEdit.Text = UCase(cboEdit.Text)\r\n'
    '      cboEdit.SelStart = pos\r\n'
    '   End If\r\n'
    'End Sub',
    content)

# ── 16: cboConsLinha_GotFocus — filtrar categorias por tipoEmpresa ─────────────
content = sub('cboConsLinha tipoEmpresa filter',
    '"SELECT DISTINCT categoria FROM produtos where (produtos.ativo = 1) ORDER BY categoria;"',
    '"SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
