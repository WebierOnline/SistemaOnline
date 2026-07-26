path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

old = '    txtISFatorConv.Text = Format(IIf(IsNull(r("fator_conversao_IS")), 0, CDbl(r("fator_conversao_IS"))), "##,##0.0000")'
new = '    txtISFatorConv.Text = Format(ValidateNull(r("fator_conversao_IS")), "##,##0.0000")'

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
