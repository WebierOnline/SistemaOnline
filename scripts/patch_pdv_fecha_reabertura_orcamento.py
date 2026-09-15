# -*- coding: utf-8 -*-
"""PDV.frm - cmdFinalizar_Click, ramo ORCAMENTO/CONSIGNADO: fecha o ciclo em Pedidos_Reabertura
ao finalizar (mesmo UPDATE que os ramos A VISTA e A PRAZO ja fazem logo apos imprimir).

Sem isso, reabrir um orcamento/consignado pra editar e resalvar deixava a linha de
Pedidos_Reabertura daquela reabertura com STATUS_PEDIDO=0 pra sempre -> a coluna "ABERTO"
em Estorno_ReabrirPedidos ficava marcada indefinidamente, mesmo ja resolvido.

Seguro pra pedido novo (nunca passou pelo Estonar): a subquery MAX(DATA)/MAX(HORA) nao acha
nenhuma linha pra esse COD_PEDIDO em Pedidos_Reabertura, UPDATE afeta 0 linhas, sem erro -
mesmo comportamento que A VISTA/A PRAZO ja tem hoje.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\PDV.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = (
    b"    End If\n"
    b"'COLOCAR O SUBTOTAL, DESCONTO E TOTAL DE CADA ITEM\n"
    b"  \n"
    b"LimparObjetos_Pedido\n"
)
n = (
    b"    End If\n"
    b"\n"
    b'dbData.Execute "UPDATE Pedidos_Reabertura SET STATUS_PEDIDO = 1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ") AND (DATA = (SELECT MAX(DATA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & "))) AND (HORA = (SELECT MAX(HORA) FROM Pedidos_Reabertura AS Pedidos_Reabertura_1 WHERE (COD_PEDIDO = " & txtCodPedido.Text & ")));"\n'
    b"\n"
    b"'COLOCAR O SUBTOTAL, DESCONTO E TOTAL DE CADA ITEM\n"
    b"  \n"
    b"LimparObjetos_Pedido\n"
)

if n in d and o not in d:
    print("[ja] cmdFinalizar_Click / ORCAMENTO-CONSIGNADO: fecha Pedidos_Reabertura")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] cmdFinalizar_Click / ORCAMENTO-CONSIGNADO: fecha Pedidos_Reabertura")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
