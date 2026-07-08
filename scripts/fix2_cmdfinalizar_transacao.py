# -*- coding: utf-8 -*-
"""
Corrige definitivamente o bug de duplicacao deixado por
fix_duplicate_transaction_blocks.py (que tambem nao avancou o ponto de
busca e por isso duplicou ainda mais o bloco A PRAZO, deixando o bloco
A VISTA sem NENHUM wrap de transacao).

1) Remove a duplicata do bloco BEGIN TRANSACTION (A PRAZO) - mantem so
   o de indentacao 8 espacos (consistente com o resto do bloco).
2) Remove a duplicata do bloco COMMIT TRANSACTION (A PRAZO).
3) Adiciona o BEGIN TRANSACTION que falta no bloco A VISTA, antes do
   comentario 'ATUALIZAR A TABELA OS (indentacao 12 espacos, igual a
   linha do comentario).
4) Adiciona o COMMIT TRANSACTION que falta no bloco A VISTA, antes do
   "If iCopiasAP <> 0 Then" final (indentacao 8 espacos, igual a linha).
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


def find_last(substr, start=0, end=None):
    end = end if end is not None else len(lines)
    found = None
    for i in range(start, end):
        if substr in lines[i]:
            found = i
    if found is None:
        raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")
    return found


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 1) Remove duplicata do BEGIN TRANSACTION (bloco A PRAZO)
# ---------------------------------------------------------------
i = find_line_exact("            lNovoCodBase = Autonumeracao_Parcelas", start, end)
assert lines[i + 1] == '            dbData.Execute "BEGIN TRANSACTION"'
assert lines[i + 2] == "            bTrans = True"
assert lines[i + 3] == "        lNovoCodBase = Autonumeracao_Parcelas"
assert lines[i + 4] == '        dbData.Execute "BEGIN TRANSACTION"'
assert lines[i + 5] == "        bTrans = True"
del lines[i : i + 3]
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 2) Remove duplicata do COMMIT TRANSACTION (bloco A PRAZO)
# ---------------------------------------------------------------
i = find_line_exact('        dbData.Execute "COMMIT TRANSACTION"', start, end)
assert lines[i + 1] == "        bTrans = False"
assert lines[i + 2] == ""
assert lines[i + 3] == '        dbData.Execute "COMMIT TRANSACTION"'
assert lines[i + 4] == "        bTrans = False"
assert lines[i + 5] == ""
del lines[i : i + 3]
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 3) Adiciona BEGIN TRANSACTION que faltava no bloco A VISTA
#    (ultima ocorrencia do comentario dentro da sub)
# ---------------------------------------------------------------
i = find_last("'ATUALIZAR A TABELA OS", start, end)
assert lines[i] == "            'ATUALIZAR A TABELA OS"
lines[i] = (
    "            lNovoCodBase = Autonumeracao_Parcelas\r\n"
    '            dbData.Execute "BEGIN TRANSACTION"\r\n'
    "            bTrans = True\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 4) Adiciona COMMIT TRANSACTION que faltava no bloco A VISTA
#    (ultima ocorrencia dentro da sub)
# ---------------------------------------------------------------
i = find_last("If iCopiasAP <> 0 Then", start, end)
assert lines[i] == "        If iCopiasAP <> 0 Then  'saber a quantidade de copias"
lines[i] = (
    '        dbData.Execute "COMMIT TRANSACTION"\r\n'
    "        bTrans = False\r\n"
    "\r\n"
    + lines[i]
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - corrigido definitivamente")
