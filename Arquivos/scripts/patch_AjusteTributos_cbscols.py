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
    c = c.replace(old, new, 1)
    print(f'{label} OK')
    return c

# 1: Cols = 17 -> 19
content = sub('Cols=19',
    '      .Cols = 17\r\n      .rows = 2\r\n      \r\n      .ColWidth(0)',
    '      .Cols = 19\r\n      .rows = 2\r\n      \r\n      .ColWidth(0)',
    content)

# 2: ColWidth — add cols 17 e 18
content = sub('ColWidth 17-18',
    '      .ColWidth(16) = 900\r\n      \r\n',
    '      .ColWidth(16) = 900\r\n'
    '      .ColWidth(17) = 900\r\n'
    '      .ColWidth(18) = 900\r\n'
    '      \r\n',
    content)

# 3: Headers — add cols 17 e 18
content = sub('Headers 17-18',
    '      .TextMatrix(0, 16) = "CEST"\r\n      \r\n',
    '      .TextMatrix(0, 16) = "CEST"\r\n'
    '      .TextMatrix(0, 17) = "CL.CBS"\r\n'
    '      .TextMatrix(0, 18) = "CL.IS"\r\n'
    '      \r\n',
    content)

# 4: Data fill — add cols 17 e 18
content = sub('DataFill 17-18',
    '            .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("var_CEST"))\r\n            \r\n',
    '            .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("var_CEST"))\r\n'
    '            .TextMatrix(.rows - 1, 17) = ValidateNull(rTabela("var_CBS"))\r\n'
    '            .TextMatrix(.rows - 1, 18) = ValidateNull(rTabela("var_IS"))\r\n'
    '            \r\n',
    content)

# 5: SELECT em cmdLocalizar_Click (FROM produtos WHERE ativo=1)
content = sub('SELECT cmdLocalizar',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST " & _\r\n'
    '      "FROM produtos " & _\r\n'
    '      "WHERE (produtos.ativo = 1) "',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS " & _\r\n'
    '      "FROM produtos " & _\r\n'
    '      "WHERE (produtos.ativo = 1) "',
    content)

# 6: SELECT em MostrarCriterios (FROM produtos var_Criterio)
content = sub('SELECT MostrarCriterios',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST " & _\r\n'
    '      "FROM produtos " & var_Criterio',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS " & _\r\n'
    '      "FROM produtos " & var_Criterio',
    content)

# 7: cmdAtualizar_Click — adiciona cClassTrib e cClassTrib_IS ao UPDATE
content = sub('cmdAtualizar CBS/IS',
    '      "cest = \'" & Grid.TextMatrix(i, 16) & "\' " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',
    '      "cest = \'" & Grid.TextMatrix(i, 16) & "\', " & _\r\n'
    '      "cClassTrib = \'" & Grid.TextMatrix(i, 17) & "\', " & _\r\n'
    '      "cClassTrib_IS = \'" & Grid.TextMatrix(i, 18) & "\' " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',
    content)

# 8: Grid_Click — permitir edicao cols 17 e 18
content = sub('GridClick 6-18',
    'For i = 6 To 16\r\n',
    'For i = 6 To 18\r\n',
    content)

# 9: txtEdit_LostFocus — adicionar casos 17 e 18
content = sub('LostFocus 17-18',
    'ElseIf iCol = 16 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n',
    'ElseIf iCol = 16 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 17 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 18 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('File written successfully.')
