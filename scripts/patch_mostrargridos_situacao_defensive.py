# -*- coding: utf-8 -*-
"""
Mesmo fix defensivo do MostrarGrid_OS, agora em MostrarGrid_OS_Situacao:
adiciona Else ao If/ElseIf de vTipoOS para nao chamar OpenRecordset com
sSQL vazio/invalido quando vTipoOS nao bate com nenhum dos 3 grupos
esperados.
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


start, end = find_sub("MostrarGrid_OS_Situacao")

marker = find_line_exact("Set r = dbData.OpenRecordset(sSQL, totalRegistros)", start, end)
outer_endif = marker - 2
assert lines[outer_endif].strip() == "End If", lines[outer_endif]

new_lines = [
    "Else",
    "    FormatarGrid_OS_Situacao Nothing",
    '    lblQuantOS.Caption = 0',
    "    Exit Sub",
]
lines[outer_endif:outer_endif] = new_lines

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
