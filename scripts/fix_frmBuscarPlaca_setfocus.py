# -*- coding: utf-8 -*-
"""
Corrige erro 5 (Invalid procedure call or argument) em frmBuscarPlaca.frm:
txtPlacaF.SetFocus nao pode rodar em Form_Load (form ainda nao esta
visivel) - move para Form_Activate.
"""

PATH = r"C:\projeto\OrdemServico\Forms\frmBuscarPlaca.frm"

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


i = find_line_exact("    txtPlacaF.SetFocus")
assert lines[i - 1] == "    ConfigurarGrid"
assert lines[i + 1] == "End Sub"
del lines[i]

i_end_load = find_line_exact("End Sub", find_line_exact("Private Sub Form_Load()"))
new_lines = [
    "",
    "Private Sub Form_Activate()",
    "    txtPlacaF.SetFocus",
    "End Sub",
]
lines[i_end_load + 1 : i_end_load + 1] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - SetFocus movido de Form_Load para Form_Activate")
