# -*- coding: utf-8 -*-
"""Estonar.frm - cmdNFCe (botao "GERAR NFCE" que o usuario acrescentou no form) so fica
ativo quando a venda NAO esta cancelada (col oculta 23) E NAO tem NFCe vinculada (col
oculta 24). Grid_Click.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = b'If Grid.TextMatrix(Grid.Row, 22) = "SIM" Or Grid.TextMatrix(Grid.Row, 23) = "SIM" Then cmdReaberturas.Enabled = True Else cmdReaberturas.Enabled = False\n'
n = (o +
     b'    If Grid.TextMatrix(Grid.Row, 23) <> "SIM" And Grid.TextMatrix(Grid.Row, 24) <> "SIM" Then cmdNFCe.Enabled = True Else cmdNFCe.Enabled = False\n')

if n in d and o not in d:
    print("[ja] cmdNFCe.Enabled")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] cmdNFCe.Enabled")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
