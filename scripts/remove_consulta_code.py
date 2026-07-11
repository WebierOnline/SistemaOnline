# -*- coding: utf-8 -*-
"""
Remove todo o codigo exclusivo da aba CONSULTA de OS_Recapadora.frm:
- as 28 Subs listadas (handlers + helpers + Preencher_*)
- a declaracao "Dim printSQL As String"
- as 6 chamadas externas a MostrarGrid_OS (em cmdAlterar_Click,
  cmdApagar_Click, cmdExcluir_Click, cmdGerarEntrada_Click,
  cmdFinalizar_Click) e o bloco de init de consulta em Form_Load
- adiciona cmdAbrirConsulta_Click no lugar de Grid_DblClick removido
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
    "MostrarGrid_OS",
    "FormatarGrid_OS",
    "Preencher_Criterios",
    "Preencher_Indice",
    "Preencher_Mostrar",
    "Preencher_TipoServico",
    "Preencher_Status",
]

removed_lines_total = 0
for name in sub_names:
    s, e = find_sub_block(name, code_start)
    n = e - s + 1
    # remove tambem uma linha em branco extra logo depois, se houver, para nao
    # acumular linhas em branco duplicadas
    end_del = e
    if e + 1 < len(lines) and lines[e + 1] == "":
        end_del = e + 1
    del lines[s : end_del + 1]
    removed_lines_total += (end_del - s + 1)
    print(f"removida sub {name}: linhas {s}-{end_del}")

# ---------------------------------------------------------------
# Grid_DblClick -> vira cmdAbrirConsulta_Click com corpo adaptado
# ---------------------------------------------------------------
s, e = find_sub_block("Grid_DblClick", code_start)
print(f"substituindo Grid_DblClick (linhas {s}-{e}) por cmdAbrirConsulta_Click")
novo_handler = [
    "Private Sub cmdAbrirConsulta_Click()",
    "OS_Consulta.lCodOSSelecionado = 0",
    "OS_Consulta.Show vbModal",
    "If OS_Consulta.lCodOSSelecionado <> 0 Then",
    "    SSTab1.Tab = 1",
    "    frmSecundario.Enabled = True",
    "    cboStatus.Enabled = True",
    "    cmdGerarEntrada.Enabled = False",
    "    cmdCancelarEntrada.Enabled = False",
    "    cmdAlterar.Enabled = True",
    "    cmdApagar.Enabled = True",
    "    cmdNovo.Enabled = True",
    '    txtCodOS.Text = ""',
    "    txtCodOS.Text = OS_Consulta.lCodOSSelecionado",
    "End If",
    "Unload OS_Consulta",
    "End Sub",
]
lines[s : e + 1] = novo_handler

# ---------------------------------------------------------------
# Dim printSQL As String
# ---------------------------------------------------------------
i = find_line_exact("Dim printSQL As String", code_start)
del lines[i]
print(f"removida linha 'Dim printSQL As String' ({i})")

# ---------------------------------------------------------------
# 6 chamadas externas a "MostrarGrid_OS" (linha solta, so a chamada,
# independente da indentacao - as 2 ocorrencias internas da aba
# CONSULTA ja sumiram junto com as subs removidas acima)
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
# bloco de init de consulta em Form_Load (Preencher_* + 5x ListIndex=0)
# ja NAO existe mais MostrarGrid_OS (removido acima). Restam:
# Preencher_TipoServico / Preencher_Mostrar / Preencher_Status /
# Preencher_Criterios / Preencher_Indice + 5 linhas ListIndex = 0
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
