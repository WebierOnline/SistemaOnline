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

# ── 1: optECest_Click → chkCest.Value = 1 ────────────────────────────────
content = sub('optECest chkCest',
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECategoria_Click()',

    '   chkCest.Value = 1\r\n'
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECategoria_Click()',
    content)

# ── 2: optECBS_Click → chkCBSIS.Value = 1 ────────────────────────────────
content = sub('optECBS chkCBSIS',
    '   lblEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optEIS_Click()',

    '   chkCBSIS.Value = 1\r\n'
    '   lblEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optEIS_Click()',
    content)

# ── 3: optClassTribCBS_Click → chkCBSIS.Value = 1 ────────────────────────
content = sub('optClassTribCBS chkCBSIS',
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',

    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()',
    content)

# ── 4: optClassTribIS_Click → chkCBSIS.Value = 1 ─────────────────────────
content = sub('optClassTribIS chkCBSIS',
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',

    'chkCBSIS.Value = 1\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
