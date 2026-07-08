# -*- coding: utf-8 -*-
"""
Aplica em cmdCancelar_Click o mesmo fix do fantasma visual do
frmVendaFechamento ja aplicado em cmdFinalizar_Click
(scripts\fix_ghost_frmvendafechamento.py): forcar uma troca real de aba
(Tab 1 e volta pra Tab 0) apos esconder o painel, para o SSTab reavaliar
a visibilidade dos seus controles corretamente.
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


start = find_line_exact("Private Sub cmdCancelar_Click()")
end = find_line_exact("End Sub", start)

i = find_line_exact("frmVendaFechamento.Visible = False", start, end)
lines[i] = (
    "frmVendaFechamento.Visible = False\r\n"
    "SSTab1.Tab = 1  'forca troca real de aba - evita frmVendaFechamento fantasma\r\n"
    "SSTab1.Tab = 0"
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
