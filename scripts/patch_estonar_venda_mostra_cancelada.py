# -*- coding: utf-8 -*-
"""Estonar.frm - pedido do usuario 2026-09-14: na consulta TipoPedido=VENDA (que compartilha
SQL com TODOS/ALUGUEL/OFICINA - "ramos A/B", chkIncompleto desmarcado), vendas canceladas
sumiam porque cmdExcluirPedido_Click APAGA as parcelas do pedido ao cancelar (`DELETE FROM
parcelas WHERE cod_pedido = ...`) e o item 5 de performance trocou o antigo `INNER JOIN
parcelas` por `EXISTS (SELECT 1 FROM parcelas WHERE ...)` - mas o efeito de excluir quem
nao tem parcela e o MESMO do INNER JOIN original (bug pre-existente, nao introduzido pela
troca pra EXISTS).

Fix 1 (ramos A e B): WHERE passa a aceitar `EXISTS(parcelas) OR pedidos.CANCELADO = 1` -
mostra a venda fechada normal (com parcelas, filtro de forma de pgto funciona) E TAMBEM a
venda cancelada (sem parcelas, sempre aparece independente do filtro de forma de pgto -
cancelada nao tem forma de pagamento pra filtrar mesmo). Ramos C1/C2 (chkIncompleto marcado)
ja mostravam canceladas hoje (FROM pedidos sem join de parcelas nenhum) - nao precisam de fix.

Fix 2 (FlexCores): linha cancelada (coluna oculta 21 = "SIM", ver patch_estonar_imgmarcada.py)
fica com o texto em vermelho (vbRed), senao preto - no mesmo loop de zebra que ja existe
(FillStyle=flexFillRepeat + ColSel, O(1) por linha, nao O(colunas)), sem loop extra.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []  # (nome, old, new, esperado)

# ---------------------------------------------------- ramos A/B: WHERE aceita cancelada sem parcela
o = b'AND EXISTS (SELECT 1 FROM parcelas WHERE parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & varTipoPgto & ") AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
n = b'AND (EXISTS (SELECT 1 FROM parcelas WHERE parcelas.COD_PEDIDO = pedidos.COD_PEDIDO " & varTipoPgto & ") OR pedidos.CANCELADO = 1) AND (pedidos.TIPO_PEDIDO <> \'ALUGUEL\')"'
reps.append(("ramos A/B: WHERE mostra cancelada mesmo sem parcelas", o, n, 2))

# ---------------------------------------------------- FlexCores: fonte vermelha se cancelada
o = (
b"Sub FlexCores(lCorPar As Long, lCorImpar As Long)\n"
b"   'ZEBRAR O FLEXGRID\n"
b"   Dim iLinha As Integer\n"
b"   Dim lCor As OLE_COLOR\n"
b"   Dim bRedrawAnt As Boolean\n"
b"   bRedrawAnt = Grid.Redraw\n"
b"   Grid.Redraw = False\n"
b"   \n"
b"   Grid.FillStyle = flexFillRepeat\n"
b"   \n"
b"   For iLinha = 1 To Grid.Rows - 1\n"
b"      With Grid\n"
b"         .Row = iLinha\n"
b"         \n"
b"         If EImpar(iLinha) Then 'Se a linha for impar:\n"
b"            lCor = lCorImpar\n"
b"         Else\n"
b"            lCor = lCorPar\n"
b"         End If\n"
b"         \n"
b"         .Col = 1                'Seleciona a partir da primeira coluna\n"
b"         .ColSel = .Cols - 1     'Seleciona at\xe9 a \xfaltima coluna\n"
b"         .CellBackColor = lCor   'Aplica a cor\n"
b"      End With\n"
b"   Next\n"
b"   \n"
b"   Grid.FillStyle = flexFillSingle\n"
b"   Grid.Redraw = bRedrawAnt\n"
b"End Sub"
)
n = (
b"Sub FlexCores(lCorPar As Long, lCorImpar As Long)\n"
b"   'ZEBRAR O FLEXGRID\n"
b"   Dim iLinha As Integer\n"
b"   Dim lCor As OLE_COLOR\n"
b"   Dim lCorTexto As OLE_COLOR\n"
b"   Dim bRedrawAnt As Boolean\n"
b"   bRedrawAnt = Grid.Redraw\n"
b"   Grid.Redraw = False\n"
b"   \n"
b"   Grid.FillStyle = flexFillRepeat\n"
b"   \n"
b"   For iLinha = 1 To Grid.Rows - 1\n"
b"      With Grid\n"
b"         .Row = iLinha\n"
b"         \n"
b"         If EImpar(iLinha) Then 'Se a linha for impar:\n"
b"            lCor = lCorImpar\n"
b"         Else\n"
b"            lCor = lCorPar\n"
b"         End If\n"
b"         \n"
b"         If .TextMatrix(iLinha, 21) = \"SIM\" Then   'coluna oculta - pedido cancelado (ver patch_estonar_imgmarcada.py)\n"
b"            lCorTexto = vbRed\n"
b"         Else\n"
b"            lCorTexto = vbBlack\n"
b"         End If\n"
b"         \n"
b"         .Col = 1                'Seleciona a partir da primeira coluna\n"
b"         .ColSel = .Cols - 1     'Seleciona at\xe9 a \xfaltima coluna\n"
b"         .CellBackColor = lCor   'Aplica a cor\n"
b"         .CellForeColor = lCorTexto   'vermelho se cancelado\n"
b"      End With\n"
b"   Next\n"
b"   \n"
b"   Grid.FillStyle = flexFillSingle\n"
b"   Grid.Redraw = bRedrawAnt\n"
b"End Sub"
)
reps.append(("FlexCores: fonte vermelha nas linhas canceladas", o, n, 1))

for nome, old, new, esperado in reps:
    if new in d and old not in d:
        print("[ja] " + nome); continue
    c = d.count(old)
    if c != esperado:
        print("[ABORTA] %s -- %d ocorrencias de old (esperava %d)" % (nome, c, esperado))
        sys.exit(1)
    if esperado > 1:
        d = d.replace(old, new)
    else:
        d = d.replace(old, new, 1)
    print("[ok]  " + nome + (" (%dx)" % c if esperado > 1 else ""))

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
