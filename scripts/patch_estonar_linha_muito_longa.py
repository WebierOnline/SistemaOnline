# -*- coding: utf-8 -*-
"""Estonar.frm - fix do erro reportado pelo usuario: linha fisica da VAZIO passou de 1023
caracteres (limite do editor/compilador do VB6) depois de eu inserir "pedidos.ValorFreteReal
as varPedFrete, pedidos.MAQUINA as varPedMaquina, " - ficou com 1087 chars. O VB6 quebrou a
linha sozinho de forma invalida ao tentar exibir/compilar (partiu a palavra "pedidos" ao
meio: "pedi _" + quebra + "s.caixa"). NENHUM outro ramo tinha esse problema (todos os outros
ja usavam `" & _` pra quebrar o SELECT em varias linhas fisicas menores - so a VAZIO e o
bloco comentado do wrapper morto de TODOS estavam inteiros numa linha so).

Fix: quebra a linha da VAZIO (codigo vivo) em 2 linhas fisicas, no mesmo ponto que
FECHADO/ABERTO/PAUSADO ja quebram (logo depois de "AS Var_NFCEInutilizada, "), com `" & _`
de verdade. Tambem quebra o comentario morto (linha 1031 chars, corrigido por seguranca -
mesmo sendo comentario, o limite de linha do VB6 e por caractere fisico, nao por token).

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []

# ---------------------------------------------------- VAZIO (codigo vivo)
o = (
b'    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, cliente.codigo, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, pedidos.ValorFreteReal as varPedFrete, pedidos.MAQUINA as varPedMaquina, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _\n'
)
n = (
b'    sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, cliente.codigo, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, " & _\n'
    b'        "pedidos.ValorFreteReal as varPedFrete, pedidos.MAQUINA as varPedMaquina, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _\n'
)
reps.append(("VAZIO: quebra linha muito longa (1087 -> 2 linhas)", o, n, 1))

# ---------------------------------------------------- comentario morto (wrapper TODOS)
o = (
b'\'sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.NFCeEnviada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, pedidos.ValorFreteReal as varPedFrete, pedidos.MAQUINA as varPedMaquina, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa " & _\n'
)
n = (
b'\'sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, cliente.nome as var_Cliente, cliente.codigo, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, pedidos.PAGAMENTO AS var_Pagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, (CASE WHEN TbNFCe.NFCeEnviada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusNFCE, (CASE WHEN TbNFCe.Inutilizada = 1 THEN \'SIM\' ELSE \'\' END) AS Var_NFCEInutilizada, " & _\n'
    b"'        \"pedidos.ValorFreteReal as varPedFrete, pedidos.MAQUINA as varPedMaquina, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa \" & _\n"
)
reps.append(("comentario morto (wrapper TODOS): quebra linha 1031 -> 2 linhas", o, n, 1))

for nome, old, new, esperado in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != esperado:
        print("[ABORTA] %s -- %d ocorrencias de old (esperava %d)" % (nome, c, esperado))
        sys.exit(1)
    d = d.replace(old, new, 1)
    print("[ok]  " + nome)

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
