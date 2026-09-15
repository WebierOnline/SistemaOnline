# -*- coding: utf-8 -*-
"""Estonar.frm - Mostrar_Pedido: corrige erro "O identificador de varias partes
'cliente.codigo' nao pode ser associado" (imagem do usuario, 2026-09-11).

Bug PRE-EXISTENTE (nao foi introduzido pelas mudancas recentes de performance/EXISTS):
`varIndice` (combo "Organizacao"/cboIndice) usa `cliente.codigo` quando cboIndice=CLIENTE,
mas varios ramos de sSQL (ex.: TODOS/VENDA, chkIncompleto desmarcado, criterio<>CLIENTE -
o "ramo A" da analise de performance) nunca fazem JOIN com a tabela Cliente -> ORDER BY
estoura pra qualquer combinacao Organizacao=CLIENTE + um desses ramos.

Fix minimo e seguro: ordenar por `pedidos.COD_CLIENTE` (coluna da propria tabela base,
sempre presente em todo FROM) em vez de `cliente.codigo` - mesma ordenacao (e a coluna
usada na condicao da join, pedidos.COD_CLIENTE = Cliente.CODIGO), sem depender de a
join existir. Nao mexe em nenhum SELECT/FROM/WHERE ja corrigido.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = (b'ElseIf cboIndice.Text = "CLIENTE" Then\n'
     b'    If cboStatus.Text = "VAZIO" Then\n'
     b'        varIndice = " pedidos.cod_pedido "\n'
     b'    Else\n'
     b'        varIndice = " cliente.codigo "\n'
     b'    End If\n')
n = (b'ElseIf cboIndice.Text = "CLIENTE" Then\n'
     b'    If cboStatus.Text = "VAZIO" Then\n'
     b'        varIndice = " pedidos.cod_pedido "\n'
     b'    Else\n'
     b"        varIndice = \" pedidos.COD_CLIENTE \"   'antes era cliente.codigo - estourava nos ramos sem JOIN Cliente (ex.: TODOS/VENDA sem criterio CLIENTE)\n"
     b'    End If\n')

c = d.count(o)
if c != 1:
    print("[ABORTA] -- %d ocorrencias de old" % c); sys.exit(1)
d = d.replace(o, n, 1)
print("[ok] varIndice CLIENTE -> pedidos.COD_CLIENTE")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
