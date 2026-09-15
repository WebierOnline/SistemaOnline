# -*- coding: utf-8 -*-
"""Estonar.frm - pedido do usuario 2026-09-14: ao imprimir a lista de vendas (cmdImprimir),
vendas canceladas mostravam a forma de pagamento vazia (sem parcela, ver
patch_estonar_venda_mostra_cancelada.py). cmdImprimir_Click reusa `printSQL` (= ultimo sSQL
de Mostrar_Pedido) direto num REL_Estorno.Relatorio.Recordset - o relatorio le var_Pagamento
pelo nome da coluna, sem passar pelo grid.

Fix (ramos A/B - unico lugar com `pgto.var_Pagamento`, distinguido dos outros 2 ramos que
tem a MESMA string por indentacao: 20 espacos = A/B, 16 espacos = CANCELADO/fora de escopo):
`(CASE WHEN pedidos.CANCELADO = 1 THEN '(cancelado)' ELSE pgto.var_Pagamento END) AS
var_Pagamento`. Como impressao e grid usam a mesma query, tambem passa a mostrar
"(cancelado)" na coluna TIPO (col 8) do grid em vez de vazio.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = b'                    "pgto.var_Pagamento,  " & _\n'
n = b'                    "(CASE WHEN pedidos.CANCELADO = 1 THEN \'(cancelado)\' ELSE pgto.var_Pagamento END) AS var_Pagamento,  " & _\n'

if n in d and o not in d:
    print("[ja] ramos A/B: var_Pagamento mostra (cancelado)")
else:
    c = d.count(o)
    if c != 2:
        print("[ABORTA] -- %d ocorrencias de old (esperava 2, so ramos A/B)" % c)
        sys.exit(1)
    d = d.replace(o, n)
    print("[ok] ramos A/B: var_Pagamento mostra (cancelado) (2x)")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
