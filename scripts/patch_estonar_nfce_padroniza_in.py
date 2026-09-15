# -*- coding: utf-8 -*-
"""Estonar.frm - Mostrar_Pedido: padroniza Var_StatusNFCE pra IN (1, 0) em todos os ramos
(pedido do usuario, 2026-09-14).

Levantamento antes de mexer: TODOS/A, TODOS/B, TODOS/C1, TODOS/C2, TODOS/CANCELADO,
TODOS/CONSIGNADO (x2), TODOS/OR�AMENTO (x2), FECHADO, ABERTO e VAZIO JA usam
`NFCeEnviada IN (1, 0)` (FECHADO/ABERTO ja convertidos no patch_estonar_perf_4e5.py, item 4).

O UNICO ramo inconsistente e o PAUSADO: seleciona `Var_StatusNFCE` DUAS vezes no mesmo
SELECT (alias duplicado) - a primeira via subquery IN(1,0), a segunda via
`TbNFCe.NFCeEnviada = 1` (joined column). ADO/DAO le sempre a PRIMEIRA coluna com esse
nome, entao a segunda definicao (`=1`) ja era morta na pratica - mas deixava o SQL confuso
e sujeito a erro se algum dia a ordem mudasse. Fix: remove a segunda definicao (`=1`) e
troca a primeira, que ainda era a forma antiga em subquery, pela forma de coluna da JOIN
que ja existe no FROM (`LEFT OUTER JOIN TbNFCe`) - mesmo padrao ja aplicado em FECHADO/ABERTO.

Ha tambem um bloco 100% COMENTADO (a linha comeca com `'sSQL = ...`, nunca executa) logo no
inicio do ramo TODOS que ainda usa `NFCeEnviada = 1` - e comentario, nao codigo, entao nao
precisa (nem faz sentido) editar.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = (
b'    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE, " & _\n'
b'        " (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.NFCeEnviada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, " & _\n'
)
n = (
b'    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, ISNULL(CASE WHEN TbNFCe.NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END, \'\') AS Var_StatusNFCE, " & _\n'
b'        " (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, pedidos.caixa as varPedCaixa, " & _\n'
)

c = d.count(o)
if c != 1:
    print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
    sys.exit(1)
d = d.replace(o, n, 1)
print("[ok] PAUSADO: remove Var_StatusNFCE duplicado (=1), padroniza pra IN (1, 0) via JOIN")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
