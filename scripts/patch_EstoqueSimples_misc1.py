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

# 1: cboDesc_Change — converter para maiusculo preservando cursor
content = sub('cboDesc_Change maiusculo',
    'Private Sub cboDesc_Change()\r\n'
    '   \'cboDesc_Click\r\n'
    'End Sub',

    'Private Sub cboDesc_Change()\r\n'
    '   Dim p As Integer\r\n'
    '   p = cboDesc.SelStart\r\n'
    '   cboDesc.Text = UCase(cboDesc.Text)\r\n'
    '   cboDesc.SelStart = p\r\n'
    'End Sub',
    content)

# 2: AvaliarFrmEdicao — optDesc passa a ativar os optE* (remover da exclusao)
content = sub('AvaliarFrmEdicao optDesc ativa',
    'If Not (optTodos.Value Or optCodBarra.Value Or optDesc.Value) Then',
    'If Not (optTodos.Value Or optCodBarra.Value) Then',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
