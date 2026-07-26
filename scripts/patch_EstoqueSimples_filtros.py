path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: optCategoria_Click — usar cboDesc em vez de cboConsLinha ───────────────
content = sub('optCategoria usa cboDesc',
    'Private Sub optCategoria_Click()\r\n'
    '   lblCategoria.Visible = True\r\n'
    '   cboConsLinha.Visible = True\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboConsLinha.SetFocus\r\n'
    'End Sub',

    'Private Sub optCategoria_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   cboConsLinha.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',
    content)

# ── 2: optDesc_Click — ocultar checkboxes ─────────────────────────────────────
content = sub('optDesc ocultar checkboxes',
    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   cboConsLinha.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   chkDescPorProduto.Visible = True\r\n'
    '   chkDescPorIniciais.Visible = True\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',

    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   cboConsLinha.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',
    content)

# ── 3: Adicionar optTags_Click e optNCM_Click apos optDesc_Click ──────────────
content = sub('add optTags + optNCM Click',
    '   cboDesc.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optMostrarFiscal_Click()',

    '   cboDesc.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optTags_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   cboConsLinha.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   cboConsLinha.Visible = False\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   chkDescPorProduto.Visible = False\r\n'
    '   chkDescPorIniciais.Visible = False\r\n'
    '   lblCodBarra.Visible = True\r\n'
    '   txtCodBarra.Visible = True\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   txtCodBarra.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optMostrarFiscal_Click()',
    content)

# ── 4: cboDesc_GotFocus — carregar conforme optCategoria/optTags ──────────────
content = sub('cboDesc_GotFocus nova logica',
    'Private Sub cboDesc_GotFocus()\r\n'
    'If chkDescPorProduto.Value = Checked Then\r\n'
    '   cboDesc.Clear\r\n'
    '   \r\n'
    '   sSQL = "SELECT DISTINCT descricao FROM produtos WHERE (produtos.ativo = 1) ORDER BY descricao;"\r\n'
    '   Set r = dbData.OpenRecordset(sSQL)\r\n'
    '   \r\n'
    '   Do While Not r.EOF\r\n'
    '      cboDesc.AddItem ValidateNull(r("descricao"))\r\n'
    '      r.MoveNext\r\n'
    '   Loop\r\n'
    '   \r\n'
    '   If r.State <> 0 Then r.Close\r\n'
    '   Set r = Nothing\r\n'
    '   \r\n'
    '   moCombo.AttachTo cboDesc\r\n'
    'End If\r\n'
    'End Sub',

    'Private Sub cboDesc_GotFocus()\r\n'
    '   cboDesc.Clear\r\n'
    '   If optCategoria.Value Then\r\n'
    '      sSQL = "SELECT DISTINCT Categoria FROM Categorias WHERE Tipo_Empresa = " & tipoEmpresa & " ORDER BY Categoria;"\r\n'
    '      Set r = dbData.OpenRecordset(sSQL)\r\n'
    '      Do While Not r.EOF\r\n'
    '         cboDesc.AddItem ValidateNull(r("Categoria"))\r\n'
    '         r.MoveNext\r\n'
    '      Loop\r\n'
    '      If r.State <> 0 Then r.Close\r\n'
    '      Set r = Nothing\r\n'
    '   ElseIf optTags.Value Then\r\n'
    '      sSQL = "SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Tipo_Empresa = " & tipoEmpresa & " ORDER BY c.Categoria, ct.Tags;"\r\n'
    '      Set r = dbData.OpenRecordset(sSQL)\r\n'
    '      Do While Not r.EOF\r\n'
    '         cboDesc.AddItem ValidateNull(r("Tags"))\r\n'
    '         r.MoveNext\r\n'
    '      Loop\r\n'
    '      If r.State <> 0 Then r.Close\r\n'
    '      Set r = Nothing\r\n'
    '   End If\r\n'
    '   moCombo.AttachTo cboDesc\r\n'
    'End Sub',
    content)

# ── 5: MostrarCriterios — optCategoria usa cboDesc + add optTags/optNCM ───────
content = sub('MostrarCriterios optCat/Tags/NCM',
    '   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = \'" & txtCodBarra.Text & "\'", "")',

    '   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = \'" & cboDesc.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = \'" & txtCodBarra.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optTags.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = \'" & cboDesc.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optNCM.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = \'" & txtCodBarra.Text & "\'", "")',
    content)

# ── 6: cmdAtualizar_Click — fix colunas (TAG deslocou tudo) + salvar TAGS ─────
content = sub('cmdAtualizar fix cols + TAGS',
    '      "descricao = \'" & TirarEspaco(Grid.TextMatrix(i, 4)) & "\', " & _\r\n'
    '      "UNID_MEDIDA = \'" & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "categoria = \'" & Grid.TextMatrix(i, 7) & "\', " & _\r\n'
    '      "fabricante = \'" & Grid.TextMatrix(i, 5) & "\', " & _\r\n'
    '      "PRATELEIRA = \'" & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 9)), ",", ".") & " " & _',

    '      "descricao = \'" & TirarEspaco(Grid.TextMatrix(i, 4)) & "\', " & _\r\n'
    '      "UNID_MEDIDA = \'" & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "categoria = \'" & Grid.TextMatrix(i, 7) & "\', " & _\r\n'
    '      "TAGS = \'" & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "fabricante = \'" & Grid.TextMatrix(i, 5) & "\', " & _\r\n'
    '      "PRATELEIRA = \'" & Grid.TextMatrix(i, 9) & "\', " & _\r\n'
    '      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 10)), ",", ".") & " " & _',
    content)

# ── 7: cmdLocalizar_Click — add optNCM focus ──────────────────────────────────
content = sub('cmdLocalizar add optNCM focus',
    'If optCodBarra.Value = True Then txtCodBarra_GotFocus\r\n'
    'End Sub',

    'If optCodBarra.Value = True Then txtCodBarra_GotFocus\r\n'
    'If optNCM.Value = True Then txtCodBarra_GotFocus\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
