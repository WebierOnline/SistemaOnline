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

# 1a: FixedCols=0 em Formatar_Grid (12 cols)
content = sub('FixedCols=0 Formatar_Grid',
    '      .Cols = 12\r\n'
    '      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 300\r\n',

    '      .Cols = 12\r\n'
    '      .rows = 2\r\n'
    '      .FixedRows = 1\r\n'
    '      .FixedCols = 0\r\n'
    '      \r\n'
    '      .ColWidth(0) = 300\r\n',
    content)

# 1b: FixedCols=0 em Formatar_Grid_Fiscal (15 cols)
content = sub('FixedCols=0 Formatar_Grid_Fiscal',
    '      .Cols = 15\r\n'
    '      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 300\r\n',

    '      .Cols = 15\r\n'
    '      .rows = 2\r\n'
    '      .FixedRows = 1\r\n'
    '      .FixedCols = 0\r\n'
    '      \r\n'
    '      .ColWidth(0) = 300\r\n',
    content)

# 2a: optCodBarra_Click — lblCodBarra caption
content = sub('optCodBarra lblCodBarra caption',
    'Private Sub optCodBarra_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Visible = True\r\n'
    '   txtCodBarra.Visible = True\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   txtCodBarra.SetFocus\r\n'
    'End Sub',

    'Private Sub optCodBarra_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Caption = "C\xf3d. Barra"\r\n'
    '   lblCodBarra.Visible = True\r\n'
    '   txtCodBarra.Visible = True\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   txtCodBarra.SetFocus\r\n'
    'End Sub',
    content)

# 2b: optDesc_Click — lblCodBarra caption + hide
content = sub('optDesc lblCodBarra caption',
    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = True\r\n'
    '   PorPalavraDupla.Visible = True\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',

    'Private Sub optDesc_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = True\r\n'
    '   PorPalavraDupla.Visible = True\r\n'
    '   lblCodBarra.Caption = "Descri\xe7\xe3o"\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',
    content)

# 2c: optCategoria_Click — lblCodBarra caption
content = sub('optCategoria lblCodBarra caption',
    'Private Sub optCategoria_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',

    'Private Sub optCategoria_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Caption = "Categoria"\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',
    content)

# 2d: optTags_Click — lblCodBarra caption
content = sub('optTags lblCodBarra caption',
    'Private Sub optTags_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',

    'Private Sub optTags_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = True\r\n'
    '   cboDesc.Visible = True\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Caption = "Tags"\r\n'
    '   lblCodBarra.Visible = False\r\n'
    '   txtCodBarra.Visible = False\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   cboDesc.SetFocus\r\n'
    'End Sub',
    content)

# 2e: optNCM_Click — lblCodBarra caption
content = sub('optNCM lblCodBarra caption',
    'Private Sub optNCM_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Visible = True\r\n'
    '   txtCodBarra.Visible = True\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   txtCodBarra.SetFocus\r\n'
    'End Sub',

    'Private Sub optNCM_Click()\r\n'
    '   lblCategoria.Visible = False\r\n'
    '   lblDesc.Visible = False\r\n'
    '   cboDesc.Visible = False\r\n'
    '   optPorPalavra.Visible = False\r\n'
    '   PorPalavraDupla.Visible = False\r\n'
    '   lblCodBarra.Caption = "NCM"\r\n'
    '   lblCodBarra.Visible = True\r\n'
    '   txtCodBarra.Visible = True\r\n'
    '   cmdLocalizar.Visible = True\r\n'
    '   txtCodBarra.SetFocus\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
