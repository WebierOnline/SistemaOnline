"""
Patch: Vendas_Consulta_PorProdutos.frm — bloco POR PRODUTOS
Adiciona filtro de tipo_pagamento no WHERE base para alinhar com
VendasServicos_Consulta, que sempre filtra 'A Vista' e 'A prazo'
mesmo quando o usuario seleciona 'TODOS'.
"""

FRM = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FRM, 'rb') as f:
    raw = f.read()

data = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('windows-1252')

old = "WHERE pedidos.CANCELADO = 0 AND pedidos.tipo_pedido <> 'ORÇAMENTO'\""
new = "WHERE pedidos.CANCELADO = 0 AND pedidos.tipo_pedido <> 'ORÇAMENTO' AND pedidos.tipo_pagamento IN ('À Vista', 'À prazo')\""

count = text.count(old)
if count != 1:
    print(f'ERRO: {count} ocorrencias (esperado 1)')
else:
    print(f'OK: encontrado')
    text = text.replace(old, new)
    result = text.encode('windows-1252').replace(b'\n', b'\r\n')
    with open(FRM, 'wb') as f:
        f.write(result)
    print('Arquivo salvo com sucesso.')
