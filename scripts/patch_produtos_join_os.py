"""
Patch: Vendas_Consulta_PorProdutos.frm — bloco POR PRODUTOS
Substitui LEFT OUTER JOIN OS por subquery correlacionada para evitar
multiplicacao de linhas quando um pedido tem mais de um registro em OS.

ANTES: LEFT OUTER JOIN OS ON pedidos.COD_PEDIDO = OS.COD_PEDIDO
       + ISNULL(OS.COD_OS, 0) AS var_CodOS no SELECT

DEPOIS: (SELECT TOP 1 COD_OS FROM OS WHERE OS.COD_PEDIDO = pedidos.COD_PEDIDO)
        diretamente no SELECT, sem JOIN
"""

FRM = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FRM, 'rb') as f:
    raw = f.read()

data = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('windows-1252')

ok = True

# Patch 1: trocar ISNULL(OS.COD_OS, 0) por subquery no SELECT
old1 = 'ISNULL(OS.COD_OS, 0) AS var_CodOS'
new1 = 'ISNULL((SELECT TOP 1 COD_OS FROM OS WHERE OS.COD_PEDIDO = pedidos.COD_PEDIDO), 0) AS var_CodOS'
c1 = text.count(old1)
if c1 != 1:
    print(f'ERRO [SELECT var_CodOS]: {c1} ocorrencias (esperado 1)')
    ok = False
else:
    print('OK [SELECT var_CodOS]')
    text = text.replace(old1, new1)

# Patch 2: remover LEFT OUTER JOIN OS do FROM
old2 = ' LEFT OUTER JOIN OS ON pedidos.COD_PEDIDO = OS.COD_PEDIDO'
new2 = ''
c2 = text.count(old2)
if c2 != 1:
    print(f'ERRO [LEFT OUTER JOIN OS]: {c2} ocorrencias (esperado 1)')
    ok = False
else:
    print('OK [LEFT OUTER JOIN OS removido]')
    text = text.replace(old2, new2)

if not ok:
    print('Erros encontrados — arquivo NAO salvo.')
else:
    result = text.encode('windows-1252').replace(b'\n', b'\r\n')
    with open(FRM, 'wb') as f:
        f.write(result)
    print('Arquivo salvo com sucesso.')
