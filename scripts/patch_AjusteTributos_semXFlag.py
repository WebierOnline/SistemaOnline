path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
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

# ── 1: Declarar bSetSemX no nivel de modulo ───────────────────────────────────
content = sub('Declarar bSetSemX',
    'Private tipoEmpresa As Long\r\n',
    'Private tipoEmpresa As Long\r\n'
    'Private bSetSemX As Boolean\r\n',
    content)

# ── 2: optCategoria_Click — set flag, remover Text= ──────────────────────────
content = sub('optCategoria flag + rm Text',
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM CATEGORIA]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optCodBarra_Click()',

    'bSetSemX = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optCodBarra_Click()',
    content)

# ── 3: optTags_Click — set flag, remover Text= ───────────────────────────────
content = sub('optTags flag + rm Text',
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM TAGS]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()',

    'bSetSemX = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()',
    content)

# ── 4: optNCM_Click — set flag, remover Text= ────────────────────────────────
content = sub('optNCM flag + rm Text',
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM NCM]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribCBS_Click()',

    'bSetSemX = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribCBS_Click()',
    content)

# ── 5: optClassTribCBS_Click — set flag, remover Text= ───────────────────────
content = sub('optClassTribCBS flag + rm Text',
    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM CBS CLASS.]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',

    'chkCBSIS.Value = 1\r\n'
    'bSetSemX = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',
    content)

# ── 6: optClassTribIS_Click — set flag, remover Text= ────────────────────────
content = sub('optClassTribIS flag + rm Text',
    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'cboConsLinha.Text = "[SEM IS CLASS.]"\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',

    'chkCBSIS.Value = 1\r\n'
    'bSetSemX = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',
    content)

# ── 7: GotFocus — definir Text=[SEM X] apos AttachTo quando flag ativo ───────
content = sub('GotFocus set Text via flag',
    '   moCombo.AttachTo cboConsLinha\r\n'
    'End Sub',

    '   moCombo.AttachTo cboConsLinha\r\n'
    '   If bSetSemX Then\r\n'
    '      bSetSemX = False\r\n'
    '      If optCategoria.Value Then\r\n'
    '         cboConsLinha.Text = "[SEM CATEGORIA]"\r\n'
    '      ElseIf optTags.Value Then\r\n'
    '         cboConsLinha.Text = "[SEM TAGS]"\r\n'
    '      ElseIf optNCM.Value Then\r\n'
    '         cboConsLinha.Text = "[SEM NCM]"\r\n'
    '      ElseIf optClassTribCBS.Value Then\r\n'
    '         cboConsLinha.Text = "[SEM CBS CLASS.]"\r\n'
    '      ElseIf optClassTribIS.Value Then\r\n'
    '         cboConsLinha.Text = "[SEM IS CLASS.]"\r\n'
    '      End If\r\n'
    '   End If\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
