path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AdicionarQuant.frm'
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

CRLF = '\r\n'
q = chr(39)  # aspas simples

# Corrigir o INSERT: '' duplo -> ' simples no CONVERT(DATETIME, ...)
content = sub('fix DATETIME quotes no INSERT',
    'CONVERT(DATETIME, ' + q + q + '" & Format(Date, ocDATA) & "' + q + q + ', 103)',
    'CONVERT(DATETIME, ' + q + '" & Format(Date, ocDATA) & "' + q + ', 103)',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
