# -*- coding: utf-8 -*-
"""
Reverte o fix anterior de cmdAlterar_Click: o usuario corrigiu a
informacao - a regra certa e "A COMECAR" oculta stProdSer, todos os
outros status (incluindo EM EXECUCAO) exibem. O codigo original ja
fazia isso; volta as 4 ocorrencias de stProdSer.Visible = False
(que eu tinha trocado por engano) de volta para True.
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


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


start, end = find_sub("cmdAlterar_Click")

count = 0
for i in range(start, end):
    if lines[i].strip() == "stProdSer.Visible = False":
        indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
        lines[i] = indent + "stProdSer.Visible = True"
        count += 1

assert count == 4, f"esperado 4 substituicoes, feito {count}"

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - revertido,", count, "ocorrencias trocadas para True")
