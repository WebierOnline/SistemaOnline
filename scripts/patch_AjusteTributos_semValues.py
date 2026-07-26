path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
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

# ── 1: cboConsLinha_GotFocus — adicionar [SEM X] na posicao 0 ────────────────
content = sub('cboConsLinha_GotFocus add SEM X',
    '   If r.State <> 0 Then r.Close\r\n'
    '   Set r = Nothing\r\n'
    '   \r\n'
    '   moCombo.AttachTo cboConsLinha\r\n'
    'End Sub',

    '   If r.State <> 0 Then r.Close\r\n'
    '   Set r = Nothing\r\n'
    '   \r\n'
    '   If optCategoria.Value Then\r\n'
    '      cboConsLinha.AddItem "[SEM CATEGORIA]", 0\r\n'
    '   ElseIf optTags.Value Then\r\n'
    '      cboConsLinha.AddItem "[SEM TAGS]", 0\r\n'
    '   ElseIf optNCM.Value Then\r\n'
    '      cboConsLinha.AddItem "[SEM NCM]", 0\r\n'
    '   ElseIf optClassTribCBS.Value Then\r\n'
    '      cboConsLinha.AddItem "[SEM CBS CLASS.]", 0\r\n'
    '   ElseIf optClassTribIS.Value Then\r\n'
    '      cboConsLinha.AddItem "[SEM IS CLASS.]", 0\r\n'
    '   End If\r\n'
    '   \r\n'
    '   moCombo.AttachTo cboConsLinha\r\n'
    'End Sub',
    content)

# ── 2: optCodBarra_Click — auto-escreve [SEM CÓD. BARRA] ─────────────────────
content = sub('optCodBarra auto SEM COD BARRA',
    'txtCodBarra.SetFocus\r\n'
    'frmEdicao.Enabled = False\r\n'
    'frmEdicaoFiltros.Enabled = False\r\n'
    'End Sub',

    'txtCodBarra.SetFocus\r\n'
    'txtCodBarra.Text = "[SEM C\xd3D. BARRA]"\r\n'
    'frmEdicao.Enabled = False\r\n'
    'frmEdicaoFiltros.Enabled = False\r\n'
    'End Sub',
    content)

# ── 3: optDesc_Click — auto-escreve [SEM DESCRIÇÃO] em cboDesc ───────────────
content = sub('optDesc auto SEM DESCRICAO',
    'cboDesc.SetFocus\r\n'
    'frmEdicao.Enabled = False\r\n'
    'frmEdicaoFiltros.Enabled = False\r\n'
    'End Sub',

    'cboDesc.SetFocus\r\n'
    'cboDesc.Text = "[SEM DESCRI\xc7\xc3O]"\r\n'
    'frmEdicao.Enabled = False\r\n'
    'frmEdicaoFiltros.Enabled = False\r\n'
    'End Sub',
    content)

# ── 4: optCategoria_Click — auto-escreve [SEM CATEGORIA] ─────────────────────
content = sub('optCategoria auto SEM CATEGORIA',
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optCodBarra_Click()',

    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM CATEGORIA]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optCodBarra_Click()',
    content)

# ── 5: optTags_Click — auto-escreve [SEM TAGS] ───────────────────────────────
content = sub('optTags auto SEM TAGS',
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()',

    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM TAGS]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()',
    content)

# ── 6: optNCM_Click — auto-escreve [SEM NCM] ─────────────────────────────────
content = sub('optNCM auto SEM NCM',
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribCBS_Click()',

    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM NCM]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribCBS_Click()',
    content)

# ── 7: optClassTribCBS_Click — auto-escreve [SEM CBS CLASS.] ─────────────────
content = sub('optClassTribCBS auto SEM CBS CLASS',
    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',

    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM CBS CLASS.]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',
    content)

# ── 8: optClassTribIS_Click — auto-escreve [SEM IS CLASS.] ───────────────────
content = sub('optClassTribIS auto SEM IS CLASS',
    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',

    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM IS CLASS.]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',
    content)

# ── 9: MostrarCriterios — 6 IIf lines -> If/Else blocks com deteccao [SEM X] ─
content = sub('MostrarCriterios SEM X where blocks',
    '   var_Criterio = var_Criterio & IIf(optCategoria.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optCodBarra.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = \'" & txtCodBarra.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optTags.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optNCM.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optClassTribCBS.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib = \'" & Left(cboConsLinha.Text, 6) & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optClassTribIS.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib_IS = \'" & Left(cboConsLinha.Text, 6) & "\'", "")',

    '   If optCategoria.Value Then\r\n'
    '      If cboConsLinha.Text = "[SEM CATEGORIA]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.categoria = \'\' OR produtos.categoria IS NULL)"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.categoria = \'" & cboConsLinha.Text & "\'"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    '   If optCodBarra.Value Then\r\n'
    '      If txtCodBarra.Text = "[SEM C\xd3D. BARRA]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cod_barra = \'\' OR produtos.cod_barra IS NULL)"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cod_barra = \'" & txtCodBarra.Text & "\'"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    '   If optTags.Value Then\r\n'
    '      If cboConsLinha.Text = "[SEM TAGS]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.TAGS = \'\' OR produtos.TAGS IS NULL)"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = \'" & cboConsLinha.Text & "\'"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    '   If optNCM.Value Then\r\n'
    '      If cboConsLinha.Text = "[SEM NCM]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.NCM = \'\' OR produtos.NCM IS NULL OR produtos.NCM = \'00000000\')"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = \'" & cboConsLinha.Text & "\'"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    '   If optClassTribCBS.Value Then\r\n'
    '      If cboConsLinha.Text = "[SEM CBS CLASS.]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cClassTrib = \'\' OR produtos.cClassTrib IS NULL)"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib = \'" & Left(cboConsLinha.Text, 6) & "\'"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    '   If optClassTribIS.Value Then\r\n'
    '      If cboConsLinha.Text = "[SEM IS CLASS.]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.cClassTrib_IS = \'\' OR produtos.cClassTrib_IS IS NULL)"\r\n'
    '      Else\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib_IS = \'" & Left(cboConsLinha.Text, 6) & "\'"\r\n'
    '      End If\r\n'
    '   End If',
    content)

# ── 10: MostrarCriterios optDesc — detectar [SEM DESCRIÇÃO] ──────────────────
content = sub('MostrarCriterios optDesc SEM DESCRICAO',
    '      Dim sPosEsp As Integer\r\n'
    '      If optPorPalavra.Value Then',

    '      Dim sPosEsp As Integer\r\n'
    '      If cboDesc.Text = "[SEM DESCRI\xc7\xc3O]" Then\r\n'
    '         var_Criterio = var_Criterio & IIf(var_Criterio <> "", " AND ", "") & "(produtos.descricao = \'\' OR produtos.descricao IS NULL)"\r\n'
    '      ElseIf optPorPalavra.Value Then',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
