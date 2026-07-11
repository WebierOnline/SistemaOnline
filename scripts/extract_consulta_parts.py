# -*- coding: utf-8 -*-
"""
Extrai (sem modificar o .frm original) os blocos de controle e as Subs
exclusivas da aba CONSULTA de OS_Recapadora.frm, salvando cada pedaco
em arquivos separados na pasta scratch, para montar OS_Consulta.frm.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"
OUT_DIR = r"C:\Users\NOTEBOOK\AppData\Local\Temp\claude\C--projeto\916fb1c0-4fd5-437b-8d03-a83de36ec5b2\scratchpad"

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


def find_control_block(control_decl_prefix, start=0):
    """Encontra 'Begin <Tipo> <Nome> ' e retorna (start_idx, end_idx_inclusive)
    usando a indentacao da propria linha Begin para achar o End correspondente."""
    i = None
    for j in range(start, len(lines)):
        if lines[j].lstrip().startswith(control_decl_prefix):
            i = j
            break
    if i is None:
        raise SystemExit(f"ERRO: controle nao encontrado: {control_decl_prefix!r}")
    indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
    end_marker = indent + "End"
    for k in range(i + 1, len(lines)):
        if lines[k] == end_marker:
            return i, k
    raise SystemExit(f"ERRO: End nao encontrado para {control_decl_prefix!r}")


def find_sub_block(sub_name, start=0):
    for i in range(start, len(lines)):
        s = lines[i].strip()
        if s == f"Private Sub {sub_name}()" or s.startswith(f"Private Sub {sub_name}("):
            for k in range(i + 1, len(lines)):
                if lines[k].strip() == "End Sub":
                    return i, k
    raise SystemExit(f"ERRO: sub nao encontrada: {sub_name!r}")


def save(name, s, e, label):
    block = "\r\n".join(lines[s : e + 1])
    fname = f"{OUT_DIR}\\consulta_{name}.txt"
    with open(fname, "w", encoding="cp1252", newline="") as f:
        f.write(block)
    print(f"{label}: linhas {s}-{e} ({e-s+1} linhas) -> {fname}")


# ---- controles ----
code_start = find_line_exact('Attribute VB_Name = "OS_Recapadora"')

s, e = find_control_block("Begin VB.Frame Frame2", 0)
save("Frame2", s, e, "Frame2")

s, e = find_control_block("Begin MSFlexGridLib.MSFlexGrid Grid ", 0)
save("Grid", s, e, "Grid")

s, e = find_control_block("Begin VB.Label lblQuant ", 0)
save("lblQuant", s, e, "lblQuant")

s, e = find_control_block("Begin VB.Label lblTotalConsulta", 0)
save("lblTotalConsulta", s, e, "lblTotalConsulta")

# ---- subs ----
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
    "Preencher_Criterios",
    "Preencher_Indice",
    "Preencher_Mostrar",
    "Preencher_Status",
    "Preencher_TipoServico",
]
for name in sub_names:
    s, e = find_sub_block(name, code_start)
    save(name, s, e, name)

# FormatarGrid_OS tem parametro, tratar separado
s, e = find_sub_block("FormatarGrid_OS", code_start)
save("FormatarGrid_OS", s, e, "FormatarGrid_OS")

print("OK - extracao concluida")
