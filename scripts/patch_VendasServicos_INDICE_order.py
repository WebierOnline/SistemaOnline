"""
patch_VendasServicos_INDICE_order.py

Substitui o ORDER BY fixo (q.var_codped) no wrapper do printSQL
pelo ORDER BY dinamico baseado na variavel INDICE.

INDICE pode conter 'pedidos_1.total' e 'pedidos_1.cod_pedido' (aliases do inner
query), entao e necessario mapear para os nomes do outer query:
  pedidos_1.total      -> var_total
  pedidos_1.cod_pedido -> var_codped
  demais valores (data_compra, nome, tipo_pagamento) ficam inalterados.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM, FRM + ".bak_INDICE_order")
print(f"Backup: {FRM}.bak_INDICE_order")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"'Envolver sSQL num subquery para adicionar vProd e vServ sem alterar cada variante de SQL\r\n"
    b"Dim sPrintBase As String\r\n"
    b"Dim idxOrd As Long\r\n"
    b"sPrintBase = sSQL\r\n"
    b"idxOrd = InStrRev(sPrintBase, \"ORDER BY\")\r\n"
    b"If idxOrd > 0 Then sPrintBase = Left(sPrintBase, idxOrd - 1)\r\n"
    b"printSQL = \"SELECT *, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(pi.total),0) FROM pedidos_itens pi WHERE pi.cod_pedido=q.var_codped) AS vProd, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(sv.total),0) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=q.var_codped) AS vServ \" & _\r\n"
    b"    \"FROM (\" & sPrintBase & \") AS q ORDER BY q.var_codped\"\r\n"
    b"End Sub\r\n"
)

NEW = (
    b"'Envolver sSQL num subquery para adicionar vProd e vServ sem alterar cada variante de SQL\r\n"
    b"Dim sPrintBase As String\r\n"
    b"Dim idxOrd As Long\r\n"
    b"Dim sOrder As String\r\n"
    b"sPrintBase = sSQL\r\n"
    b"idxOrd = InStrRev(sPrintBase, \"ORDER BY\")\r\n"
    b"If idxOrd > 0 Then sPrintBase = Left(sPrintBase, idxOrd - 1)\r\n"
    b"sOrder = Replace(Replace(INDICE, \";\", \"\"), \"pedidos_1.\", \"\")\r\n"
    b"sOrder = Replace(sOrder, \"total\", \"var_total\")\r\n"
    b"sOrder = Replace(sOrder, \"cod_pedido\", \"var_codped\")\r\n"
    b"printSQL = \"SELECT *, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(pi.total),0) FROM pedidos_itens pi WHERE pi.cod_pedido=q.var_codped) AS vProd, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(sv.total),0) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=q.var_codped) AS vServ \" & _\r\n"
    b"    \"FROM (\" & sPrintBase & \") AS q ORDER BY \" & sOrder\r\n"
    b"End Sub\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: ORDER BY do printSQL agora usa variavel INDICE")
print("Arquivo salvo com sucesso.")
