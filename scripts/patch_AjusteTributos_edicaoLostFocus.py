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

# ── 1: Limpar controles apos cmdEdicaoColetiva ────────────────────────────
content = sub('clear after cmdEdicaoColetiva',
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub',

    'picAguarde.Visible = False\r\n'
    'txtEdicaoColetiva.Text = ""\r\n'
    'cboEdicaoColetiva.Text = ""\r\n'
    'ResetarMarcas\r\n'
    'End Sub',
    content)

# ── 2: Novos handlers LostFocus ───────────────────────────────────────────
new_handlers = (
    'Private Sub txtEdicaoColetiva_LostFocus()\r\n'
    'Dim s As String\r\n'
    's = Trim(txtEdicaoColetiva.Text)\r\n'
    'If optENCM.Value Or optECest.Value Then\r\n'
    '   s = Replace(s, ".", "")\r\n'
    '   s = Replace(s, " ", "")\r\n'
    'End If\r\n'
    'txtEdicaoColetiva.Text = s\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdicaoColetiva_LostFocus()\r\n'
    '   cboEdicaoColetiva.Text = Trim(cboEdicaoColetiva.Text)\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('Insert LostFocus handlers',
    'Private Sub ResetarMarcas()',
    new_handlers + 'Private Sub ResetarMarcas()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
