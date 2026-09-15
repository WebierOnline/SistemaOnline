# -*- coding: utf-8 -*-
"""Estorno_ReabrirPedidos.frm - loadInformacoes: pedido do usuario 2026-09-15 (opcao 1,
"mais simples"). ABERTO passava a marcar SO reabertura ainda pendente (status_pedido=0);
usuario testou e viu que uma reabertura ja resolvida (finalizada de novo) ficava com
CANCEL. e ABERTO os dois em branco - queria ver a marcacao mesmo assim.

Fix: ABERTO vira o espelho de CANCEL. - marca toda linha que NAO for cancelamento
(`cancelado = 0`), resolvida ou nao. Cada linha do historico vira sempre OU cancelamento
OU reabertura, nunca as duas em branco. Perde a distincao visual "essa reabertura especifica
ainda ta pendente" (info ainda existe no pedidos.status_pedido/reaberto, so nao aparece mais
nessa tela) - decisao explicita do usuario, escolheu a opcao simples entre as 2 oferecidas.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estorno_ReabrirPedidos.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = b"CASE status_pedido WHEN 0 THEN 'SIM' ELSE '' END AS vStatus"
n = b"CASE cancelado WHEN 0 THEN 'SIM' ELSE '' END AS vStatus"

if n in d and o not in d:
    print("[ja] ABERTO = espelho de CANCEL.")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] ABERTO = espelho de CANCEL. (cancelado=0)")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
