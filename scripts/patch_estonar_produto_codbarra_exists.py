# -*- coding: utf-8 -*-
"""Estonar.frm - Mostrar_Pedido: corrige os criterios PRODUTO e C\xd3D. BARRA.

Nenhuma das 13 variantes de sSQL faz JOIN com pedidos_itens, mas varCriterio (montado
1x so, antes do bloco de status, e concatenado no fim via
`sSQL = sSQL & "" & varCriterio & " ORDER BY " & varIndice` - vale pra TODOS os ramos)
referenciava `pedidos_itens.cod_produto` direto -> SQL Server estourava "multi-part
identifier could not be bound" em 100% das combinacoes de status/tipo.

Fix: EXISTS correlacionado por pedidos.COD_PEDIDO (mesmo padrao do item 5 de performance -
evita JOIN que poderia duplicar linha se o produto aparecer mais de uma vez no mesmo pedido).
`pedidos.COD_PEDIDO` sempre existe (toda ramo tem FROM pedidos sem alias).

O caso "0" (busca vazia proposital) do C\xd3D. BARRA usava um truque esquisito
(`pedidos_itens.cod_produto < '0'`) que tambem nunca rodou de verdade (mesmo bug); trocado
pelo sentinela padrao ja usado no resto do arquivo: `pedidos.cod_pedido = 0`.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []

# ---------------------------------------------------- C\xd3D. BARRA
o = (b'ElseIf cboCriterios.Text = "C\xd3D. BARRA" Then\n'
     b'    If txtCodProdutoBarra.Text = "0" Then\n'
     b"        varCriterio = \" and (pedidos_itens.cod_produto < '0') \"\n"
     b'    Else\n'
     b'        varCriterio = " and (pedidos_itens.cod_produto = " & txtCodProdutoBarra.Text & ") "\n'
     b'    End If\n')
n = (b'ElseIf cboCriterios.Text = "C\xd3D. BARRA" Then\n'
     b'    If txtCodProdutoBarra.Text = "0" Then\n'
     b'        varCriterio = " and pedidos.cod_pedido = 0 "\n'
     b'    Else\n'
     b'        varCriterio = " and EXISTS (SELECT 1 FROM pedidos_itens WHERE pedidos_itens.cod_pedido = pedidos.COD_PEDIDO AND pedidos_itens.cod_produto = " & txtCodProdutoBarra.Text & ") "\n'
     b'    End If\n')
reps.append(("C\xd3D. BARRA -> EXISTS pedidos_itens", o, n))

# ---------------------------------------------------- PRODUTO
o = (b'ElseIf cboCriterios.Text = "PRODUTO" Then\n'
     b'    If txtCodProduto.Text = "" Then Exit Sub\n'
     b'    varCriterio = " and (pedidos_itens.cod_produto = " & txtCodProduto.Text & ")"\n'
     b'End If\n')
n = (b'ElseIf cboCriterios.Text = "PRODUTO" Then\n'
     b'    If txtCodProduto.Text = "" Then Exit Sub\n'
     b'    varCriterio = " and EXISTS (SELECT 1 FROM pedidos_itens WHERE pedidos_itens.cod_pedido = pedidos.COD_PEDIDO AND pedidos_itens.cod_produto = " & txtCodProduto.Text & ")"\n'
     b'End If\n')
reps.append(("PRODUTO -> EXISTS pedidos_itens", o, n))

for nome, old, new in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != 1:
        print("[ABORTA] %s -- %d ocorrencias de old" % (nome, c)); sys.exit(1)
    d = d.replace(old, new, 1)
    print("[ok]  " + nome)

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
