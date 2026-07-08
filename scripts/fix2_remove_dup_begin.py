# -*- coding: utf-8 -*-
"""Remove a duplicata do bloco BEGIN TRANSACTION no bloco A PRAZO."""
PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()
text = raw.decode("cp1252")
lines = text.split("\r\n")

OLD = [
    "            lNovoCodBase = Autonumeracao_Parcelas",
    '            dbData.Execute "BEGIN TRANSACTION"',
    "            bTrans = True",
    "        lNovoCodBase = Autonumeracao_Parcelas",
    '        dbData.Execute "BEGIN TRANSACTION"',
    "        bTrans = True",
    "        'ATUALIZAR A TABELA OS",
]
NEW = [
    "        lNovoCodBase = Autonumeracao_Parcelas",
    '        dbData.Execute "BEGIN TRANSACTION"',
    "        bTrans = True",
    "        'ATUALIZAR A TABELA OS",
]

joined_old = "\r\n".join(OLD)
joined_new = "\r\n".join(NEW)
text2 = "\r\n".join(lines)
n = text2.count(joined_old)
assert n == 1, f"esperado 1, achou {n}"
text2 = text2.replace(joined_old, joined_new, 1)

out = text2.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)
print("OK")
