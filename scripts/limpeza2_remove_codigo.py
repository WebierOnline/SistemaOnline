# -*- coding: utf-8 -*-
"""
Limpeza (2a vez) da aba CONSULTA em OS_Recapadora.frm:
Remove todo o codigo exclusivo da aba CONSULTA:
- as 28 Subs listadas (handlers + helpers + Preencher_*)
- a declaracao "Dim printSQL As String"
- as 6 chamadas externas a MostrarGrid_OS e o bloco de init em Form_Load
- Grid_DblClick e removido por completo (nao convertido em botao -
  Menu_Consulta_OS_Click ja existe e cumpre esse papel)
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


def find_sub_block(sub_name, start=0):
    for i in range(start, len(lines)):
        s = lines[i].strip()
        if s == f"Private Sub {sub_name}()" or s.startswith(f"Private Sub {sub_name}("):
            for k in range(i + 1, len(lines)):
                if lines[k].strip() == "End Sub":
                    return i, k
    raise SystemExit(f"ERRO: sub nao encontrada: {sub_name!r}")


code_start = find_line_exact('Attribute VB_Name = "OS_Recapadora"')

sub_names = [
    "cboConsultaCriterios_Click",
    "cboConsultaCriterios_GotFocus",
    "cboConsultaCriterios_Validate",
    "cboConsultaMostrar_Change",
    "cboConsultaMostrar_Click",
    "cboConsultaMostrar_GotFocus",
    "cboConsultaMostrar_Validate",
    "cboConsultaStatus_Change",
    "cboConsultaStatus_Click",
    "cboConsultaStatus_GotFocus",
    "cboConsultaStatus_Validate",
    "cboIndice_Change",
    "cboIndice_Click",
    "cboIndice_GotFocus",
    "cboLocalizar_GotFocus",
    "cboLocalizar_LostFocus",
    "cboTipoServico_Change",
    "cboTipoServico_Click",
    "cboTipoServico_GotFocus",
    "cmdExibir_Click",
    "cmdImprimirConsulta_Click",
    "Grid_DblClick",
    "MostrarGrid_OS",
    "FormatarGrid_OS",
    "Preencher_Criterios",
    "Preencher_Indice",
    "Preencher_Mostrar",
    "Preencher_Status",
    "Preencher_TipoServico",
]

for name in sub_names:
    s, e = find_sub_block(name, code_start)
    n = e - s + 1
    end_del = e
    if e + 1 < len(lines) and lines[e + 1] == "":
        end_del = e + 1
    del lines[s : end_del + 1]
    print(f"removida sub {name}: linhas {s}-{end_del}")

# ---------------------------------------------------------------
# Dim printSQL As String
# ---------------------------------------------------------------
i = find_line_exact("Dim printSQL As String", code_start)
del lines[i]
print(f"removida linha 'Dim printSQL As String' ({i})")

# ---------------------------------------------------------------
# 6 chamadas externas a "MostrarGrid_OS" (linha solta)
# ---------------------------------------------------------------
removidas = 0
i = code_start
while i < len(lines):
    if lines[i].strip() == "MostrarGrid_OS":
        del lines[i]
        removidas += 1
    else:
        i += 1
print(f"removidas {removidas} chamadas soltas a MostrarGrid_OS")
assert removidas == 6, f"esperado 6 chamadas externas, removidas {removidas}"

# ---------------------------------------------------------------
# bloco de init de consulta em Form_Load
# ---------------------------------------------------------------
i_start = find_line_exact("Preencher_TipoServico", code_start)
expected = [
    "Preencher_TipoServico",
    "Preencher_Mostrar",
    "Preencher_Status",
    "Preencher_Criterios",
    "Preencher_Indice",
    "cboConsultaMostrar.ListIndex = 0",
    "cboConsultaStatus.ListIndex = 0",
    "cboConsultaCriterios.ListIndex = 0",
    "cboTipoServico.ListIndex = 0",
    "cboIndice.ListIndex = 0",
]
actual = lines[i_start : i_start + len(expected)]
assert actual == expected, f"bloco de init nao bate:\n{actual}\nvs\n{expected}"
del lines[i_start : i_start + len(expected)]
print(f"removido bloco de init de consulta em Form_Load (linhas {i_start}-{i_start+len(expected)-1})")

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - codigo de consulta removido")
