# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - optDataPrevisao:
1) AtualizarCamposCriterios: mostra/esconde junto com optDataEntrada/
   optDataTermino nos criterios DATA/PERIODO/MENSAL.
2) MostrarGrid_OS: campoData trata optDataPrevisao igual optDataTermino
   (consulta continua filtrando por os.DATA_TERMINO).
3) FormatarGrid_OS: coluna TERMINO (2) - quando optDataPrevisao marcado
   e OS.STATUS <> 'TERMINADO', mostra a DATA_TERMINO (previsao) com
   sufixo "(P)"; quando STATUS = 'TERMINADO', mostra a data normal
   (sem sufixo) independente da opcao marcada; sem optDataPrevisao e
   status <> TERMINADO, continua vazio (comportamento anterior).
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

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


# ---------------------------------------------------------------
# 1) AtualizarCamposCriterios
# ---------------------------------------------------------------
i = find_line_exact("Private Sub AtualizarCamposCriterios()")
end = find_line_exact("End Sub", i)

j = find_line_exact("optDataTermino.Visible = False", i, end)
lines[j] = "optDataTermino.Visible = False\r\noptDataPrevisao.Visible = False"
end = find_line_exact("End Sub", i)

count = 0
k = i
while True:
    found = None
    for idx in range(i, end):
        if lines[idx].strip() == "optDataTermino.Visible = True":
            found = idx
            break
    if found is None:
        break
    lines[found] = lines[found] + "\r\n   optDataPrevisao.Visible = True"
    count += 1
    end = find_line_exact("End Sub", i)
assert count == 3, count

# ---------------------------------------------------------------
# 2) MostrarGrid_OS: campoData - NAO precisa mudar. optDataEntrada/
#    optDataTermino/optDataPrevisao sao mutuamente exclusivos (mesmo
#    grupo de OptionButton), entao "Else -> os.DATA_TERMINO" ja cobre
#    optDataPrevisao automaticamente (so verificado, nao alterado).
# ---------------------------------------------------------------

# ---------------------------------------------------------------
# 3) FormatarGrid_OS: coluna TERMINO com sufixo "(P)"
# ---------------------------------------------------------------
s_line = find_line_exact('         If IsNull(rTabela("DATA_TERMINO")) Or rTabela("var_status") <> "TERMINADO" Then')
e_line = find_line_exact("         End If", s_line, s_line + 5)
old_block = lines[s_line : e_line + 1]
expected = [
    '         If IsNull(rTabela("DATA_TERMINO")) Or rTabela("var_status") <> "TERMINADO" Then',
    '            .TextMatrix(.Rows - 1, 2) = ""',
    "         Else",
    '            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy")',
    "         End If",
]
assert old_block == expected, old_block

novo_block = [
    '         If IsNull(rTabela("DATA_TERMINO")) Then',
    '            .TextMatrix(.Rows - 1, 2) = ""',
    '         ElseIf rTabela("var_status") = "TERMINADO" Then',
    '            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy")',
    "         ElseIf optDataPrevisao.Value = True Then",
    '            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy") & "(P)"',
    "         Else",
    '            .TextMatrix(.Rows - 1, 2) = ""',
    "         End If",
]
lines[s_line : e_line + 1] = novo_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
