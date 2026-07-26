path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

old = (
    '            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("vAtivo"))' + R +
    '            ' + R +
    '            rTabela.MoveNext'
)

new = (
    '            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("vAtivo"))' + R +
    '            .Row = .rows - 1' + R +
    '            .Col = 0' + R +
    '            .ColSel = .Cols - 1' + R +
    '            .FillStyle = 1' + R +
    '            If .TextMatrix(.rows - 1, 12) = "DESATIVO" Then' + R +
    '                .CellForeColor = RGB(139, 0, 0)' + R +
    '            Else' + R +
    '                .CellForeColor = vbBlack' + R +
    '            End If' + R +
    '            .FillStyle = 0' + R +
    '            ' + R +
    '            rTabela.MoveNext'
)

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
