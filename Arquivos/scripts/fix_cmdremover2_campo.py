path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# cmdAdicionar2 grava txtQuantFinal em produtos_entrada_itens.QuantidadeTributavel (nao em QUANT)
# portanto a leitura no cmdRemover2 deve usar QuantidadeTributavel, nao QUANT

old = (
    '   dQtdRem = CDbl(SQLExecutaRetorno("SELECT ISNULL(QUANT,0) r FROM produtos_entrada_itens WHERE codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1), "r", "0"))\r\n'
)
new = (
    '   dQtdRem = CDbl(SQLExecutaRetorno("SELECT ISNULL(QuantidadeTributavel,0) r FROM produtos_entrada_itens WHERE codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1), "r", "0"))\r\n'
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
