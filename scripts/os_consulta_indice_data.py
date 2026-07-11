# -*- coding: utf-8 -*-
"""
OS_Consulta.frm:
1) Preencher_Indice: remove "TIPO DE SERVIÇO".
2) MostrarGrid_OS: move o calculo de campoData pra ANTES do calculo de
   INDICE (estava depois), e o ramo "DATA" do INDICE passa a usar
   campoData (respeitando optDataEntrada/optDataTermino/optDataPrevisao)
   em vez do fixo os.DATA_ENTRADA. Remove tambem o ramo morto
   "TIPO DE SERVIÇO" do INDICE (item removido do combo).
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
# 1) Preencher_Indice
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Preencher_Indice()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub Preencher_Indice()",
    "   cboIndice.Clear",
    '   cboIndice.AddItem "CÓD. OS"',
    '   cboIndice.AddItem "TIPO DE SERVIÇO"',
    '   cboIndice.AddItem "CLIENTE"',
    '   cboIndice.AddItem "DATA"',
    "End Sub",
]
assert old == expected, old
novo = [
    "Private Sub Preencher_Indice()",
    "   cboIndice.Clear",
    '   cboIndice.AddItem "CÓD. OS"',
    '   cboIndice.AddItem "CLIENTE"',
    '   cboIndice.AddItem "DATA"',
    "End Sub",
]
lines[i : end + 1] = novo

# ---------------------------------------------------------------
# 2) MostrarGrid_OS: mover campoData pra antes do INDICE + usar no ramo DATA
# ---------------------------------------------------------------
i = find_line_exact("Private Sub MostrarGrid_OS()")
end = find_line_exact("End Sub", i)

# remove o bloco campoData da posicao atual (mais abaixo)
j = find_line_exact("'campo de data usado nos filtros DATA/PERÍODO/MENSAL", i, end)
campo_data_block = lines[j : j + 7]
expected_campo = [
    "'campo de data usado nos filtros DATA/PERÍODO/MENSAL",
    "Dim campoData As String",
    "If optDataEntrada.Value = True Then",
    '   campoData = "os.DATA_ENTRADA"',
    "Else",
    '   campoData = "os.DATA_TERMINO"',
    "End If",
]
assert campo_data_block == expected_campo, campo_data_block
del lines[j : j + 7]
end = find_line_exact("End Sub", i)

# remove o ramo "TIPO DE SERVIÇO" do INDICE e troca o ramo "DATA"
k = find_line_exact("'indice", i, end)
old_indice_block = lines[k : k + 12]
expected_indice = [
    "'indice",
    'If cboIndice.Text = "CÓD. OS" Then',
    '   INDICE = "os.COD_OS DESC "',
    'ElseIf cboIndice.Text = "TIPO DE SERVIÇO" Then',
    '   INDICE = "os.TIPO_OS DESC "',
    'ElseIf cboIndice.Text = "CLIENTE" Then',
    '   INDICE = "cliente.nome DESC "',
    'ElseIf cboIndice.Text = "DATA" Then',
    '   INDICE = "os.DATA_ENTRADA DESC "',
    "Else",
    '   INDICE = "OS.COD_OS DESC "',
    "End If",
]
assert old_indice_block == expected_indice, old_indice_block

novo_indice_block = campo_data_block + [
    "",
    "'indice",
    'If cboIndice.Text = "CÓD. OS" Then',
    '   INDICE = "os.COD_OS DESC "',
    'ElseIf cboIndice.Text = "CLIENTE" Then',
    '   INDICE = "cliente.nome DESC "',
    'ElseIf cboIndice.Text = "DATA" Then',
    '   INDICE = campoData & " DESC "',
    "Else",
    '   INDICE = "OS.COD_OS DESC "',
    "End If",
]
lines[k : k + 12] = novo_indice_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
