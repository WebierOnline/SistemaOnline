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

content = sub('Foco cboEdicaoColetiva para CBS/IS',
    '   If Not bNumOk Then\r\n'
    '      MsgBox "Digite exatamente " & iExact & " d\xedgitos num\xe9ricos.", vbExclamation, "Edi\xe7\xe3o Coletiva"\r\n'
    '      txtEdicaoColetiva.SetFocus\r\n'
    '      Exit Sub\r\n'
    '   End If',

    '   If Not bNumOk Then\r\n'
    '      MsgBox "Digite exatamente " & iExact & " d\xedgitos num\xe9ricos.", vbExclamation, "Edi\xe7\xe3o Coletiva"\r\n'
    '      If optECBS.Value Or optEIS.Value Then\r\n'
    '         cboEdicaoColetiva.SetFocus\r\n'
    '      Else\r\n'
    '         txtEdicaoColetiva.SetFocus\r\n'
    '      End If\r\n'
    '      Exit Sub\r\n'
    '   End If',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
