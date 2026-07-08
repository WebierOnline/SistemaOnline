"""
Patch: Vendas_Consulta_PorProdutos.frm — bloco POR PRODUTOS
Troca o filtro do WHERE base:
  ANTES: pedidos_itens.cancelado = 0 AND pedidos.tipo_pedido <> 'ORCAMENTO'
  DEPOIS: pedidos.CANCELADO = 0 AND pedidos.tipo_pedido <> 'ORCAMENTO'

Por que: VendasServicos_Consulta exclui pedidos cancelados (CANCELADO=0)
e soma TODOS os itens do pedido (inclusive itens cancelados).
POR PRODUTOS deve seguir a mesma logica para que os totais batam.
"""

FRM = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FRM, 'rb') as f:
    raw = f.read()

data = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('windows-1252')

old = "WHERE pedidos_itens.cancelado = 0 AND pedidos.tipo_pedido <> 'ORÇAMENTO'"
new = "WHERE pedidos.CANCELADO = 0 AND pedidos.tipo_pedido <> 'ORÇAMENTO'"

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
