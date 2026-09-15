# -*- coding: utf-8 -*-
"""Estonar.frm - pedido do usuario 2026-09-15: ao clicar numa linha de venda cancelada,
so cmdMostrarProdutos (EXIBIR PRODUTOS) e cmdImprimir (IMPRIMIR LISTA) devem ficar ativos.

Mapeamento dos botoes (TX = legenda real do chameleonButton, "Caption" no .frm nao e o texto
visivel - vem tudo "00"):
  cmdPedidoAbrir=REABRIR, cmdModificar/cmdModificarConsignado=EDITAR, cmdExcluirPedido=CANCELAR,
  cmdPedidoImprimir=REIMPRIMIR, cmdPDF=CRIAR PDF, cmdMostrarProdutos=EXIBIR PRODUTOS,
  cmdImprimir=IMPRIMIR LISTA, cmdReaberturas=REABERTURAS.
  (cmdReimprimir e um botao "&Abrir" vestigial, Enabled nunca tocado em lugar nenhum do
  codigo - nao faz parte do conjunto de acoes de linha, nao mexi nele.)

Achado: `LiberarBotoesPermissoes` (chamada no fim de Grid_Click) reabilita cmdPedidoAbrir/
cmdExcluirPedido/cmdModificar/cmdModificarConsignado com base SO na permissao do usuario
logado, sem considerar o status da linha - por isso hoje uma venda cancelada deixava tudo
liberado igual uma venda normal.

Fix: apos LiberarBotoesPermissoes (pra ter a ultima palavra), se a linha clicada for
cancelada (coluna oculta 23 = "SIM", ver patch_estonar_renumerar_colunas.py) forca
cmdPedidoAbrir/cmdModificar/cmdModificarConsignado/cmdExcluirPedido/cmdPedidoImprimir/
cmdPDF/cmdReaberturas = False. cmdMostrarProdutos e cmdImprimir ficam como a logica ja
deixou (True).

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = (
    b"    'permiss\xf5es\n"
    b"    LiberarBotoesPermissoes\n"
    b"End If\n"
    b"End Sub"
)
n = (
    b"    'permiss\xf5es\n"
    b"    LiberarBotoesPermissoes\n"
    b"    \n"
    b"    'venda cancelada: so pode ver os produtos e imprimir a lista\n"
    b'    If Grid.TextMatrix(Grid.Row, 23) = "SIM" Then\n'
    b"        cmdPedidoAbrir.Enabled = False\n"
    b"        cmdModificar.Enabled = False\n"
    b"        cmdModificarConsignado.Enabled = False\n"
    b"        cmdExcluirPedido.Enabled = False\n"
    b"        cmdPedidoImprimir.Enabled = False\n"
    b"        cmdPDF.Enabled = False\n"
    b"        cmdReaberturas.Enabled = False\n"
    b"    End If\n"
    b"End If\n"
    b"End Sub"
)

if n in d and o not in d:
    print("[ja] Grid_Click: restricao de botoes pra venda cancelada")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] Grid_Click: restricao de botoes pra venda cancelada")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
