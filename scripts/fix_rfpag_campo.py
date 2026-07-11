# -*- coding: utf-8 -*-
"""
Corrige o Campo de rfPag em REL_OS_Consulta.frm: desfaz a edicao errada
feita via Edit tool (escreveu o literal "\\xe1" em vez do acento) e
grava corretamente em binario cp1252: "= P\xe1g.: [Pagina] de [Paginas]"
"""

PATH = r"C:\projeto\Compartilhado\Forms\REL_OS_Consulta.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")

bad = '         Campo           =   "= P\\xe1g.: [Pagina] de [Paginas]"'
assert text.count(bad) == 1, text.count(bad)

good = '         Campo           =   "= P\xe1g.: [Pagina] de [Paginas]"'
text = text.replace(bad, good)

text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(text.encode("cp1252"))

print("OK - Campo do rfPag corrigido")
