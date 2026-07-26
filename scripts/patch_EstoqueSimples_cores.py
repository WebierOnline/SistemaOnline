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

# Adicionar loop de cores alternadas (branco/cinza) em Formatar_Grid e Formatar_Grid_Fiscal
content = sub('cores alternadas grid',
    '      .rows = .rows - 1\r\n'
    '      .Redraw = True\r\n'
    '      picAguarde.Visible = False\r\n'
    '   End With',

    '      .rows = .rows - 1\r\n'
    '      Dim lRow As Long\r\n'
    '      .FillStyle = 1\r\n'
    '      For lRow = 1 To .rows - 1\r\n'
    '         .Row = lRow\r\n'
    '         .Col = 0\r\n'
    '         .ColSel = .Cols - 1\r\n'
    '         If lRow Mod 2 = 0 Then\r\n'
    '            .CellBackColor = &HE0E0E0\r\n'
    '         Else\r\n'
    '            .CellBackColor = vbWhite\r\n'
    '         End If\r\n'
    '      Next lRow\r\n'
    '      .FillStyle = 0\r\n'
    '      .Redraw = True\r\n'
    '      picAguarde.Visible = False\r\n'
    '   End With',
    content, replace_all=True)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
