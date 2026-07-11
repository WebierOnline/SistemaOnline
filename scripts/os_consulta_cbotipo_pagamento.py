# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - cboTipo (TODOS/À VISTA/À PRAZO), filtrando
os.tipo_pagamento em MostrarGrid_OS e MostrarGrid_OS_Refinado.
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
# 1) Preencher_Tipo + cboTipo_GotFocus - inseridos apos Preencher_TipoServico
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Preencher_TipoServico()")
end = find_line_exact("End Sub", i)

novos_subs = [
    "",
    "Private Sub Preencher_Tipo()",
    "cboTipo.Clear",
    'cboTipo.AddItem "TODOS"',
    'cboTipo.AddItem "À VISTA"',
    'cboTipo.AddItem "À PRAZO"',
    "End Sub",
    "",
    "Private Sub cboTipo_GotFocus()",
    "Dim itemAtual As String",
    "itemAtual = cboTipo.Text",
    "Preencher_Tipo",
    "cboTipo.Text = itemAtual",
    "moCombo.AttachTo cboTipo",
    "End Sub",
]
lines[end + 1 : end + 1] = novos_subs

# ---------------------------------------------------------------
# 2) Form_Load: popular e selecionar TODOS
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Form_Load()")
end = find_line_exact("End Sub", i)
j = find_line_exact("cboTipoServico.ListIndex = 0", i, end)
lines[j] = 'cboTipoServico.ListIndex = 0\r\nPreencher_Tipo\r\ncboTipo.Text = "TODOS"'

# ---------------------------------------------------------------
# 3) MostrarGrid_OS: varTipoPagamento computado + injetado antes do
#    ORDER BY nos 18 ramos
# ---------------------------------------------------------------
i = find_line_exact("Private Sub MostrarGrid_OS()")
end = find_line_exact("End Sub", i)

j = find_line_exact("'campo de data usado nos filtros DATA/PERÍODO/MENSAL", i, end)
novo_calc = [
    "'forma de pagamento (TODOS/À VISTA/À PRAZO)",
    "Dim varTipoPagamento As String",
    'If cboTipo.Text = "À VISTA" Then',
    '   varTipoPagamento = "AND (os.tipo_pagamento = \'À Vista\') "',
    'ElseIf cboTipo.Text = "À PRAZO" Then',
    '   varTipoPagamento = "AND (os.tipo_pagamento = \'À Prazo\') "',
    "Else",
    '   varTipoPagamento = ""',
    "End If",
    "",
]
lines[j:j] = novo_calc
end = find_line_exact("End Sub", i)

count = 0
k = i
while True:
    k = None
    for idx in range(i, end):
        if lines[idx].strip() == '"ORDER BY " & INDICE':
            k = idx
            break
    if k is None:
        break
    lines[k] = lines[k].replace(
        '"ORDER BY " & INDICE', 'varTipoPagamento & "ORDER BY " & INDICE'
    )
    count += 1
    end = find_line_exact("End Sub", i)
assert count == 18, count

# ---------------------------------------------------------------
# 4) MostrarGrid_OS_Refinado: mesmo calculo + injetado antes do ORDER BY
# ---------------------------------------------------------------
i = find_line_exact("Private Sub MostrarGrid_OS_Refinado()")
end = find_line_exact("End Sub", i)

j = find_line_exact('If txtFiltroRefinado.Text = "" Then', i, end)
k = find_line_exact("End If", j, j + 4)
novo_calc2 = [
    "",
    "'forma de pagamento (TODOS/À VISTA/À PRAZO)",
    "Dim varTipoPagamento As String",
    'If cboTipo.Text = "À VISTA" Then',
    '   varTipoPagamento = "AND (os.tipo_pagamento = \'À Vista\') "',
    'ElseIf cboTipo.Text = "À PRAZO" Then',
    '   varTipoPagamento = "AND (os.tipo_pagamento = \'À Prazo\') "',
    "Else",
    '   varTipoPagamento = ""',
    "End If",
]
lines[k + 1 : k + 1] = novo_calc2
end = find_line_exact("End Sub", i)

m = find_line_exact('   "ORDER BY os.DATA_TERMINO DESC"', i, end)
lines[m] = '   varTipoPagamento & "ORDER BY os.DATA_TERMINO DESC"'

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK -", count, "ramos de MostrarGrid_OS atualizados + Refinado")
