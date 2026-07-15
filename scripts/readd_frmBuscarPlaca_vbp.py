# -*- coding: utf-8 -*-
"""Re-adiciona Form=Forms\\frmBuscarPlaca.frm ao .vbp (removido no teste A/B)."""

PATH = r"C:\projeto\OrdemServico\OrdemServico.vbp"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")

anchor = "Form=Forms\\OS_Consulta.frm"
n = lines.count(anchor)
assert n == 1, n
idx = lines.index(anchor)
lines.insert(idx + 1, "Form=Forms\\frmBuscarPlaca.frm")

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - frmBuscarPlaca.frm re-registrado no .vbp")
