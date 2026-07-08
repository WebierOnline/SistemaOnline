# -*- coding: utf-8 -*-
"""
Corrige o "fantasma" visual do frmVendaFechamento (aba FINANCEIRO, Tab
index 2) que ficava sobreposto na aba SITUACAO apos cmdFinalizar_Click
dar certo.

Causa: frmVendaFechamento pertence ao Tab(2) do SSTab1 (Tab(2).Control(0)
= "frmVendaFechamento"). O SSTab so reavalia/repinta corretamente a
visibilidade dos controles de cada aba quando ocorre uma troca REAL de
aba (evento Click). Como o codigo so fazia "frmVendaFechamento.Visible =
False" manualmente e em seguida "SSTab1.Tab = 0" (valor que ja era 0,
sem troca real), o SSTab nunca reprocessava aquela area - so sumia se o
usuario clicasse em outra aba e voltasse (troca real).

Fix: forcar uma troca real de aba (Tab 1 e volta para Tab 0) em vez de
so atribuir Tab = 0 direto, reproduzindo em codigo o mesmo workaround
manual que o usuario ja confirmou que funciona.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

i = find_line_exact("SSTab1.Tab = 0", start, end)
assert lines[i - 1] == "MostrarGrid_OS_Situacao"
lines[i] = (
    "SSTab1.Tab = 1  'forca troca real de aba - evita frmVendaFechamento fantasma\r\n"
    "SSTab1.Tab = 0"
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
