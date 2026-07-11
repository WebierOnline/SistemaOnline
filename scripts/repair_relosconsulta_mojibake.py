# -*- coding: utf-8 -*-
"""
Repara corrupcao de encoding introduzida por um uso indevido do Edit tool
(UTF-8) num arquivo .frm que e' cp1252 - 5 ocorrencias do padrao
"\xef\xbf\xbd" (replacement char gravado como bytes UTF-8) em 4 linhas.
"""

PATH = r"C:\projeto\Compartilhado\Forms\REL_OS_Consulta.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")

pairs = [
    ('   Caption         =   "RELAT\xef\xbf\xbdRIO DE ORDEM DE SERVI\xef\xbf\xbdOS"',
     '   Caption         =   "RELAT\xd3RIO DE ORDEM DE SERVI\xc7OS"'),
    ('         Caption         =   "CRIT\xef\xbf\xbdRIO:"',
     '         Caption         =   "CRIT\xc9RIO:"'),
    ('         Caption         =   "T\xef\xbf\xbdCNICO"',
     '         Caption         =   "T\xc9CNICO"'),
    ('         Caption         =   "RELAT\xef\xbf\xbdRIO DE CONTAS"',
     '         Caption         =   "RELAT\xd3RIO DE CONTAS"'),
]

for bad, good in pairs:
    n = text.count(bad)
    assert n == 1, (n, bad)
    text = text.replace(bad, good)

pat = chr(0xEF) + chr(0xBF) + chr(0xBD)
assert text.count(pat) == 0, text.count(pat)

text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(text.encode("cp1252"))

print("OK - 5 ocorrencias de mojibake reparadas")
