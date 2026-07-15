# -*- coding: utf-8 -*-
"""
TESTE A/B: remove a linha "Form=Forms\\frmBuscarPlaca.frm" do .vbp
(o arquivo .frm continua no disco, so deixa de estar registrado no
projeto) para confirmar se a mera presenca no .vbp e' o gatilho do
erro 713/440 no Start.
"""

PATH = r"C:\projeto\OrdemServico\OrdemServico.vbp"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")

target = "Form=Forms\\frmBuscarPlaca.frm"
n = lines.count(target)
assert n == 1, n
lines.remove(target)

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - linha removida do .vbp (teste A/B)")
