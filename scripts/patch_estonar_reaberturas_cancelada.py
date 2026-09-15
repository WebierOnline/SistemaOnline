# -*- coding: utf-8 -*-
"""Estonar.frm - pedido do usuario 2026-09-15: cmdReaberturas deve ficar ativo tambem quando
o pedido foi CANCELADO (mesmo sem nunca ter sido reaberto) - a tabela Pedidos_Reabertura
tambem guarda o evento de cancelamento (CANCELADO=1), entao ha historico pra mostrar.

3 pontos:
1. Grid_Click: habilita cmdReaberturas se Reaberto(22)="SIM" OU Cancelado(23)="SIM".
2. Grid_Click: tira cmdReaberturas do bloco que forca tudo desabilitado em venda cancelada
   (patch_estonar_botoes_cancelada.py) - ele agora e uma excecao, fica ativo.
3. cmdReaberturas_Click: so bloqueia (MsgBox "sem historico") se NEM reaberto NEM cancelado.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

reps = []

# 1) Grid_Click - condicao de habilitar
o1 = b'If Grid.TextMatrix(Grid.Row, 22) = "SIM" Then cmdReaberturas.Enabled = True Else cmdReaberturas.Enabled = False\n'
n1 = b'If Grid.TextMatrix(Grid.Row, 22) = "SIM" Or Grid.TextMatrix(Grid.Row, 23) = "SIM" Then cmdReaberturas.Enabled = True Else cmdReaberturas.Enabled = False\n'
reps.append(("Grid_Click: habilita cmdReaberturas tambem se cancelada", o1, n1, 1))

# 2) Grid_Click - remove cmdReaberturas do bloco de venda cancelada
o2 = (
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
)
n2 = (
    b"    'venda cancelada: so pode ver os produtos, imprimir a lista e ver o historico (cmdReaberturas)\n"
    b'    If Grid.TextMatrix(Grid.Row, 23) = "SIM" Then\n'
    b"        cmdPedidoAbrir.Enabled = False\n"
    b"        cmdModificar.Enabled = False\n"
    b"        cmdModificarConsignado.Enabled = False\n"
    b"        cmdExcluirPedido.Enabled = False\n"
    b"        cmdPedidoImprimir.Enabled = False\n"
    b"        cmdPDF.Enabled = False\n"
    b"    End If\n"
)
reps.append(("Grid_Click: cmdReaberturas vira excecao no bloco cancelada", o2, n2, 1))

# 3) cmdReaberturas_Click - so bloqueia se NEM reaberto NEM cancelado
o3 = b'If Grid.TextMatrix(Grid.Row, 22) <> "SIM" Then\n    MsgBox "N\xe3o existe um hist\xf3rico de reabertura para esse pedido!", vbInformation, "Aviso do Sistema"\n    Exit Sub\nEnd If\n'
n3 = b'If Grid.TextMatrix(Grid.Row, 22) <> "SIM" And Grid.TextMatrix(Grid.Row, 23) <> "SIM" Then\n    MsgBox "N\xe3o existe um hist\xf3rico de reabertura ou cancelamento para esse pedido!", vbInformation, "Aviso do Sistema"\n    Exit Sub\nEnd If\n'
reps.append(("cmdReaberturas_Click: guard aceita reaberto OU cancelado", o3, n3, 1))

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
