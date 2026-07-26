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

content = sub('optEIS chkCBSIS',
    '   cboEdicaoColetiva.Visible = True\r\n'
    '   cmdEdicaoColetiva.Visible = True\r\n'
    '   lblEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cmdEdicaoColetiva_Click()',

    '   chkCBSIS.Value = 1\r\n'
    '   cboEdicaoColetiva.Visible = True\r\n'
    '   cmdEdicaoColetiva.Visible = True\r\n'
    '   lblEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cmdEdicaoColetiva_Click()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
