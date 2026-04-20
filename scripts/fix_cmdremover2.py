path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# Substituir UPDATE quant_estoque que usava col 6 (agora UnidadeTributavel)
# pela quantidade correta lida da tabela (produtos_entrada_itens.QUANT)
# e usar Val() para evitar problema com formato "0000" na col 4 (CodigoProduto)

old = (
    '   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 6) & " WHERE (codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 4) & ");"\r\n'
)
new = (
    '   Dim dQtdRem As Double\r\n'
    '   dQtdRem = CDbl(SQLExecutaRetorno("SELECT ISNULL(QUANT,0) r FROM produtos_entrada_itens WHERE codigo = " & GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 1), "r", "0"))\r\n'
    '   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(dQtdRem, ",", ".") & " WHERE (codigo = " & Val(GridEntradaItens.TextMatrix(GridEntradaItens.RowSel, 4)) & ");"\r\n'
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
