# -*- coding: utf-8 -*-
"""
Limpeza (2a vez) da aba CONSULTA em OS_Recapadora.frm:
Remove os controles Frame2, Grid (bare), lblQuant, lblTotalConsulta.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_control_block(control_decl_prefix, start=0):
    i = None
    for j in range(start, len(lines)):
        if lines[j].lstrip().startswith(control_decl_prefix):
            i = j
            break
    if i is None:
        raise SystemExit(f"ERRO: controle nao encontrado: {control_decl_prefix!r}")
    indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
    end_marker = indent + "End"
    for k in range(i + 1, len(lines)):
        if lines[k] == end_marker:
            return i, k
    raise SystemExit(f"ERRO: End nao encontrado para {control_decl_prefix!r}")


begins_before = sum(1 for l in lines if l.strip().startswith("Begin "))
ends_before = sum(1 for l in lines if l.strip() == "End")

removed = 0
for prefix in [
    "Begin VB.Frame Frame2",
    "Begin MSFlexGridLib.MSFlexGrid Grid ",
    "Begin VB.Label lblQuant ",
    "Begin VB.Label lblTotalConsulta",
]:
    s, e = find_control_block(prefix, 0)
    n = e - s + 1
    del lines[s : e + 1]
    removed += n
    print(f"removido {prefix!r}: linhas {s}-{e} ({n} linhas)")

begins_after = sum(1 for l in lines if l.strip().startswith("Begin "))
ends_after = sum(1 for l in lines if l.strip() == "End")

assert (begins_before - begins_after) == (ends_before - ends_after), (
    begins_before, begins_after, ends_before, ends_after
)
print("Begin antes/depois:", begins_before, begins_after)
print("End(bare) antes/depois:", ends_before, ends_after)

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - controles removidos")
