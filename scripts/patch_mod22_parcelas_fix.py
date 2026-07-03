# -*- coding: utf-8 -*-
"""
Fix: query de display do PRODUTO em loadPedidos (10 espacos antes de "FROM).
O script anterior patcheou apenas a query de count (6 espacos).
Quando totalRegistros >= 1, sSQL e' reatribuida a esta query sem var_CodBarra/var_CodProd,
causando erro "item nao encontrado" em FormatarGrid_Itens.
"""
import sys

FILE = r'C:\Projeto\Compartilhado\Forms\Parcelas_Consulta_Produtos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

old = (
    "pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS \" & _\n"
    "          \"FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto \" & _\n"
)
new = (
    "pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS, ISNULL(produtos.COD_BARRA,'') as var_CodBarra, pedidos_itens.cod_produto as var_CodProd \" & _\n"
    "          \"FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto \" & _\n"
)

count = text.count(old)
if count != 1:
    print(f'ERRO: {count} ocorrencias (esperado 1)')
    sys.exit(1)
text = text.replace(old, new)
print(f'OK (1x): SQL PRODUTO display query — adiciona var_CodBarra e var_CodProd')

out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
