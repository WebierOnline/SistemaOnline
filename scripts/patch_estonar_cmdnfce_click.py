# -*- coding: utf-8 -*-
"""Estonar.frm - novo Sub cmdNFCe_Click: gera e imprime a NFCe do pedido selecionado
chamando PDV.GerarEImprimirNFCeParaPedido (patch_pdv_gerar_nfce_pedido.py). Guards batem
com a condicao de habilitar o botao (patch_estonar_cmdnfce_enable ja aplicado em Grid_Click):
col 23 = cancelado, col 24 = ja tem NFCe.

Screen.MousePointer = vbHourglass durante a chamada - TransmitirNFCe fala com a SEFAZ pela
rede, pode demorar alguns segundos. Mostrar_Pedido no final atualiza o grid (coluna NFCe)
sem precisar clicar Exibir de novo.

.frm cp1252 -> edicao binaria + normaliza CRLF.
"""
import sys

p = r"C:\projeto\PDV\Forms\Estonar.frm"
d = open(p, "rb").read().replace(b"\r\n", b"\n").replace(b"\r", b"\n")

o = b'Private Sub cmdReaberturas_Click()\n'
n = (
b"Private Sub cmdNFCe_Click()\n"
b"If Grid.Rows <= 1 Then\n"
b'    MsgBox "N\xe3o existe nenhum pedido selecionado!", vbInformation, "Aviso do Sistema"\n'
b"    Exit Sub\n"
b"End If\n"
b"\n"
b'If Grid.TextMatrix(Grid.Row, 23) = "SIM" Then\n'
b'    MsgBox "N\xe3o \xe9 poss\xedvel gerar NFCe de um pedido cancelado!", vbInformation, "Aviso do Sistema"\n'
b"    Exit Sub\n"
b"End If\n"
b"\n"
b'If Grid.TextMatrix(Grid.Row, 24) = "SIM" Then\n'
b'    MsgBox "Esse pedido j\xe1 possui uma NFCe vinculada!", vbInformation, "Aviso do Sistema"\n'
b"    Exit Sub\n"
b"End If\n"
b"\n"
b"Screen.MousePointer = vbHourglass\n"
b"PDV.GerarEImprimirNFCeParaPedido CLng(Grid.TextMatrix(Grid.Row, 2))\n"
b"Screen.MousePointer = vbDefault\n"
b"\n"
b"Mostrar_Pedido\n"
b"End Sub\n"
b"\n"
b"Private Sub cmdReaberturas_Click()\n"
)

if n in d and o not in d:
    print("[ja] cmdNFCe_Click")
else:
    c = d.count(o)
    if c != 1:
        print("[ABORTA] -- %d ocorrencias de old (esperava 1)" % c)
        sys.exit(1)
    d = d.replace(o, n, 1)
    print("[ok] cmdNFCe_Click adicionado")

open(p, "wb").write(d.replace(b"\n", b"\r\n"))
print("gravado " + p)
