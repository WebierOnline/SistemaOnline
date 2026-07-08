# -*- coding: utf-8 -*-
"""
Restaura a adicao de cod_mecanico na primeira metade do UNION de
MostrarGrid_Servicos, que foi perdida pelo script de reparo de encoding
(a linha continha acento corrompido + minha adicao; o reparo restaurou
so o acento a partir do HEAD, descartando minha adicao de cod_mecanico).
Usa uma ancora ASCII-only para nao mexer no trecho acentuado.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

OLD = "CODIGO AS var_CODITEM, '' as var_CODPROD FROM OS_Servicos_Auto WHERE (COD_OS = "
NEW = "CODIGO AS var_CODITEM, '' as var_CODPROD, cod_mecanico FROM OS_Servicos_Auto WHERE (COD_OS = "

n = text.count(OLD)
assert n == 1, f"esperado 1 ocorrencia, encontrado {n}"
text = text.replace(OLD, NEW, 1)

out = text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - restaurado")
