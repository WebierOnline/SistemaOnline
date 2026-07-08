# -*- coding: utf-8 -*-
"""
Corrige bug do patch anterior: os passos 9 e 11 (bloco A VISTA) nao
avancaram o ponto de busca, entao encontraram de novo a MESMA primeira
ocorrencia de "'ATUALIZAR A TABELA OS" / "If iCopiasAP <> 0 Then" (bloco
A PRAZO) em vez da segunda ocorrencia (bloco A VISTA) - duplicando
BEGIN/COMMIT TRANSACTION no bloco A PRAZO e deixando o bloco A VISTA sem
nenhum dos dois.

1) Remove a duplicata de BEGIN TRANSACTION (bloco A PRAZO).
2) Remove a duplicata de COMMIT TRANSACTION (bloco A PRAZO).
3) Adiciona lNovoCodBase + BEGIN TRANSACTION antes de
   "'ATUALIZAR A TABELA OS" do bloco A VISTA.
4) Adiciona COMMIT TRANSACTION antes de "If iCopiasAP <> 0 Then" do
   bloco A VISTA.
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


def find_line(substr, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if substr in lines[i]:
            return i
    raise SystemExit(f"ERRO: ancora nao encontrada: {substr!r}")


start = find_line_exact("Private Sub cmdFinalizar_Click()")
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 1) Remove duplicata do BEGIN TRANSACTION (bloco A PRAZO)
# ---------------------------------------------------------------
i = find_line_exact("            lNovoCodBase = Autonumeracao_Parcelas", start, end)
assert lines[i + 1] == '            dbData.Execute "BEGIN TRANSACTION"'
assert lines[i + 2] == "            bTrans = True"
j = find_line_exact("        lNovoCodBase = Autonumeracao_Parcelas", start, end)
assert lines[j + 1] == '        dbData.Execute "BEGIN TRANSACTION"'
assert lines[j + 2] == "        bTrans = True"
assert j == i + 3, (i, j)
del lines[j : j + 3]
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 2) Remove duplicata do COMMIT TRANSACTION (bloco A PRAZO)
# ---------------------------------------------------------------
i = find_line_exact('        dbData.Execute "COMMIT TRANSACTION"', start, end)
assert lines[i + 1] == "        bTrans = False"
j = find_line_exact('        dbData.Execute "COMMIT TRANSACTION"', i + 1, end)
assert lines[j + 1] == "        bTrans = False"
assert lines[j - 1] == ""  # linha em branco entre os dois blocos
del lines[j - 1 : j + 2]
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 3) Adiciona BEGIN TRANSACTION que faltava no bloco A VISTA
# ---------------------------------------------------------------
i = find_line("'ATUALIZAR A TABELA OS", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = (
    indent + "lNovoCodBase = Autonumeracao_Parcelas\r\n"
    + indent + 'dbData.Execute "BEGIN TRANSACTION"\r\n'
    + indent + "bTrans = True\r\n"
    + lines[i]
)
end = find_line_exact("End Sub", start)

# ---------------------------------------------------------------
# 4) Adiciona COMMIT TRANSACTION que faltava no bloco A VISTA
# ---------------------------------------------------------------
i = find_line("If iCopiasAP <> 0 Then", start, end)
indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
lines[i] = (
    indent + 'dbData.Execute "COMMIT TRANSACTION"\r\n'
    + indent + "bTrans = False\r\n\r\n"
    + lines[i]
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - corrigido")
