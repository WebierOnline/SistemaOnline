# -*- coding: utf-8 -*-
"""Estonar.frm - Mostrar_Pedido, itens 4+5 da ordem de ataque de performance.

Escopo SEGURO (preserva o resultado de cada ramo, so mexe onde a join ja existe
ou onde a conversao e mecanica 1:1):

  Item 5 (matar DISTINCT + blowup de parcelas):
    - ramo A  (TODOS/VENDA/ALUGUEL/OFICINA, chkIncompleto desmarcado, criterio<>CLIENTE)
    - ramo B  (idem, criterio=CLIENTE)
      -> tira `INNER JOIN parcelas` (so servia pro filtro cboTipoPgto),
         troca por `EXISTS (SELECT 1 FROM parcelas WHERE ... <varTipoPgto>)`,
         move varTipoPgto do WHERE externo pra dentro do EXISTS,
         remove `SELECT DISTINCT` -> `SELECT` (sem a join 1:N nao ha o que deduplicar;
         pgto e 1:1 via GROUP BY, cliente e 1:1 via PK).

  Item 4 (subquery correlacionada -> coluna da join que ja existe):
    - ramo FECHADO e ramo ABERTO ja tem `LEFT OUTER JOIN TbNFCe ON Num_OS_VD_Origem = COD_PEDIDO`
      -> `ISNULL((SELECT CASE WHEN N.NFCeEnviada IN (1,0) ... FROM TbNFCe AS N WHERE ...), '')`
         vira `ISNULL(CASE WHEN TbNFCe.NFCeEnviada IN (1,0) ... END, '')` (mesma predicado, mesmo ISNULL).
    - no ramo B, como a FROM ja esta sendo reescrita e o `INNER JOIN cliente` continua,
      a subquery de var_Cliente vira `cliente.nome AS var_Cliente`.

NAO tocado de proposito (anotado pro usuario):
    - PAUSADO: seleciona Var_StatusNFCE DUAS vezes (subquery IN(1,0) + TbNFCe.NFCeEnviada=1);
      remover a subquery mudaria a logica exibida -> fica fora.
    - VAZIO: `INNER JOIN Cliente` + `COD_CLIENTE IS NULL` => query retorna 0 linhas sempre
      (bug pre-existente) -> otimizar nao muda nada.
    - subqueries de var_Cliente nos ramos sem join de cliente (A, C1, C2, CANC, CONSIG/naoCLI,
      ORCAM/naoCLI, CONSIG/CLI, ORCAM/CLI): ganho marginal (lookup por PK), fica pra depois.
    - ramos fora de A/B continuam quebrados quando cboTipoPgto <> TODOS (parcelas.FORMA_PGTO
      sem parcelas no FROM) -> bug pre-existente, fora do escopo de 4+5.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []

# ----------------------------------------------------------------- RAMO A
A_old = (
b'                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _\n'
b'                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _\n'
b'                    "pgto.var_Pagamento,  " & _\n'
b'                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE " & _\n'
b'                    "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO " & _\n'
b'                    "LEFT JOIN (SELECT pp.COD_PEDIDO, SUBSTRING((SELECT \', \' + p2.FORMA_PGTO FROM dbo.parcelas p2 WHERE p2.COD_PEDIDO = pp.COD_PEDIDO FOR XML PATH (\'\')), 2, 1000) AS var_Pagamento FROM dbo.parcelas pp GROUP BY pp.COD_PEDIDO) pgto ON pgto.COD_PEDIDO = pedidos.COD_PEDIDO " & _\n'
b'                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
)
A_new = (
b'                sSQL = "SELECT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _\n'
b'                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _\n'
b'                    "pgto.var_Pagamento,  " & _\n'
b'                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE " & _\n'
b'                    "FROM pedidos " & _\n'
b'                    "LEFT JOIN (SELECT pp.COD_PEDIDO, SUBSTRING((SELECT \', \' + p2.FORMA_PGTO FROM dbo.parcelas p2 WHERE p2.COD_PEDIDO = pp.COD_PEDIDO FOR XML PATH (\'\')), 2, 1000) AS var_Pagamento FROM dbo.parcelas pp GROUP BY pp.COD_PEDIDO) pgto ON pgto.COD_PEDIDO = pedidos.COD_PEDIDO " & _\n'
b'                    "WHERE " & varStatus & " " & varFormaPgto & " " & vTipoPedido & " AND EXISTS (SELECT 1 FROM parcelas WHERE parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & varTipoPgto & ") AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
)
reps.append(("A - matar DISTINCT + parcelas->EXISTS", A_old, A_new))

# ----------------------------------------------------------------- RAMO B
B_old = (
b'                sSQL = "SELECT DISTINCT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _\n'
b'                    "(SELECT nome AS var_Cliente FROM Cliente AS C WHERE (c.CODIGO = pedidos.COD_CLIENTE)) AS var_Cliente, " & _\n'
b'                    "pgto.var_Pagamento,  " & _\n'
b'                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE " & _\n'
b'                    "FROM pedidos INNER JOIN parcelas ON pedidos.COD_PEDIDO = parcelas.COD_PEDIDO INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _\n'
b'                    "LEFT JOIN (SELECT pp.COD_PEDIDO, SUBSTRING((SELECT \', \' + p2.FORMA_PGTO FROM dbo.parcelas p2 WHERE p2.COD_PEDIDO = pp.COD_PEDIDO FOR XML PATH (\'\')), 2, 1000) AS var_Pagamento FROM dbo.parcelas pp GROUP BY pp.COD_PEDIDO) pgto ON pgto.COD_PEDIDO = pedidos.COD_PEDIDO " & _\n'
b'                    "WHERE " & varStatus & " " & varFormaPgto & " " & varTipoPgto & "" & vTipoPedido & "  AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
)
B_new = (
b'                sSQL = "SELECT pedidos.cod_pedido AS var_CodPedido, pedidos.TIPO_PEDIDO AS var_TIPOPedido, pedidos.DATA_COMPRA as var_Data, pedidos.SUBTOTAL as var_Subtotal, pedidos.ValorDescReal as var_Desc, pedidos.ValorAcrescReal as var_Acresc, pedidos.TOTAL var_Total, pedidos.COD_FUNCIONARIO as varCod_Func, pedidos.TIPO_PEDIDO AS var_TipoPedido, pedidos.TIPO_PAGAMENTO AS var_TipoPagamento, (CASE WHEN pedidos.status_pedido = 1 THEN \'FECHADO\' ELSE \'ABERTO\' END) AS Var_StatusPedido, (CASE WHEN pedidos.reaberto = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusREABERTO, (CASE WHEN pedidos.CANCELADO = 1 THEN \'SIM\' ELSE \'\' END) AS Var_StatusCANCELADO, pedidos.caixa as varPedCaixa, pedidos.codcaixa as varPedCodCaixa,  " & _\n'
b'                    "cliente.nome AS var_Cliente, " & _\n'
b'                    "pgto.var_Pagamento,  " & _\n'
b'                    "ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE " & _\n'
b'                    "FROM pedidos INNER JOIN cliente ON pedidos.COD_CLIENTE = cliente.CODIGO " & _\n'
b'                    "LEFT JOIN (SELECT pp.COD_PEDIDO, SUBSTRING((SELECT \', \' + p2.FORMA_PGTO FROM dbo.parcelas p2 WHERE p2.COD_PEDIDO = pp.COD_PEDIDO FOR XML PATH (\'\')), 2, 1000) AS var_Pagamento FROM dbo.parcelas pp GROUP BY pp.COD_PEDIDO) pgto ON pgto.COD_PEDIDO = pedidos.COD_PEDIDO " & _\n'
b'                    "WHERE " & varStatus & " " & varFormaPgto & " " & vTipoPedido & " AND EXISTS (SELECT 1 FROM parcelas WHERE parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & varTipoPgto & ") AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
)
reps.append(("B - matar DISTINCT + parcelas->EXISTS + var_Cliente vira coluna da join", B_old, B_new))

# ----------------------------------------------------------------- RAMO FECHADO
F_old = b'Var_StatusCANCELADO, ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE, " & _\n'
F_new = b'Var_StatusCANCELADO, ISNULL(CASE WHEN TbNFCe.NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END, \'\') AS Var_StatusNFCE, " & _\n'
reps.append(("FECHADO - Var_StatusNFCE subquery -> coluna do LEFT JOIN TbNFCe", F_old, F_new))

# ----------------------------------------------------------------- RAMO ABERTO
AB_old = b'Var_StatusCANCELADO , ISNULL ((SELECT (CASE WHEN N .NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END) FROM TbNFCe AS N WHERE (Num_OS_VD_Origem = pedidos.COD_PEDIDO)), \'\') AS Var_StatusNFCE, " & _\n'
AB_new = b'Var_StatusCANCELADO , ISNULL(CASE WHEN TbNFCe.NFCeEnviada IN (1, 0) THEN \'SIM\' ELSE \'\' END, \'\') AS Var_StatusNFCE, " & _\n'
reps.append(("ABERTO - Var_StatusNFCE subquery -> coluna do LEFT JOIN TbNFCe", AB_old, AB_new))

# ----------------------------------------------------------------- aplica
for nome, old, new in reps:
    if new in d and old not in d:
        print("[ja] " + nome)
        continue
    c = d.count(old)
    if c != 1:
        print("[ABORTA] %s -- %d ocorrencias de old" % (nome, c))
        sys.exit(1)
    d = d.replace(old, new, 1)
    print("[ok]  " + nome)

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
