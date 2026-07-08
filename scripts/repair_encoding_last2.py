# -*- coding: utf-8 -*-
"""Corrige as ultimas 2 linhas com corrupcao de encoding em OS_Recapadora.frm."""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

TARGET = chr(0xEF) + chr(0xBF) + chr(0xBD)

OLD1 = 'Caption         =   "Respons' + TARGET + 'vel"'
NEW1 = 'Caption         =   "Responsável"'
assert text.count(OLD1) == 1, text.count(OLD1)
text = text.replace(OLD1, NEW1, 1)

OLD2 = 'TX              =   "Or' + TARGET + 'amento PDF"'
NEW2 = 'TX              =   "Orçamento PDF"'
assert text.count(OLD2) == 1, text.count(OLD2)
text = text.replace(OLD2, NEW2, 1)

remaining = text.count(TARGET)
print("restantes apos essa passada:", remaining)

out = text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("saved, bytes:", len(out))
