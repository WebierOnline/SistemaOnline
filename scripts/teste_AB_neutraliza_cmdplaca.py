# -*- coding: utf-8 -*-
"""
TESTE A/B temporario: esvazia o corpo de cmdPlaca_Click (remove a
referencia a frmBuscarPlaca) para isolar a causa do erro 713/440.
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


start = find_line_exact("Private Sub cmdPlaca_Click()")
end = find_line_exact("End Sub", start, start + 25)

new_lines = [
    "Private Sub cmdPlaca_Click()",
    "'TESTE AB - corpo temporariamente vazio",
    "End Sub",
]

lines[start : end + 1] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - cmdPlaca_Click esvaziado (linhas {start}-{end} substituidas)")
