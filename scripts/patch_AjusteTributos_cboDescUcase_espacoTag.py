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

# ── 1: Permitir espaco no cboEdicaoColetiva_KeyPress para optETags ────────
content = sub('KeyPress espaco para Tags',
    '      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then\r\n'
    '         KeyAscii = 0\r\n'
    '      End If',

    '      If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 32) Then\r\n'
    '         KeyAscii = 0\r\n'
    '      End If',
    content)

# ── 2: cboDesc_Change — adicionar logica uppercase ────────────────────────
content = sub('cboDesc_Change uppercase',
    'Private Sub cboDesc_Change()\r\n'
    '   \'cboDesc_Click\r\n'
    'End Sub',

    'Private Sub cboDesc_Change()\r\n'
    '   Dim pos As Integer\r\n'
    '   pos = cboDesc.SelStart\r\n'
    '   cboDesc.Text = UCase(cboDesc.Text)\r\n'
    '   cboDesc.SelStart = pos\r\n'
    'End Sub',
    content)

# ── 3: cboDesc_KeyPress — converter minuscula para maiuscula ──────────────
content = sub('cboDesc_KeyPress uppercase',
    'Private Sub cboDesc_LostFocus()\r\n'
    '   \'cboDesc_Click\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboConsLinha_Click()',

    'Private Sub cboDesc_KeyPress(KeyAscii As Integer)\r\n'
    '   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboDesc_LostFocus()\r\n'
    '   \'cboDesc_Click\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboConsLinha_Click()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
