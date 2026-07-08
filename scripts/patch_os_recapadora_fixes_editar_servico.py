# -*- coding: utf-8 -*-
"""
Patch OS_Recapadora.frm - 3 ajustes reportados pelo usuario:

1) cmdNovo_Click nao resetava o estado dos botoes de servico (podiam ficar
   presos em "modo edicao" de uma OS anterior). Agora forca Adicionar/Remover
   habilitados, Editar desabilitado, e limpa a selecao de mecanico/edicao.
2) Grid_Servicos_DblClick nao preenchia cboServicosAuto (so setava vServico
   direto, sem tocar em cboServicosAuto/txtCodServicoAuto para nao disparar
   o side-effect de txtCodServicoAuto_Change sobrescrever o preco). Agora
   tambem seta cboServicosAuto.Text, que nao dispara esse Change (soh LostFocus
   disparado por foco real faz isso).
3) cmdEditarOS_Click: ao reabrir a OS, se o status for "EM EXECUÇÃO",
   stProdSer deve ficar invisivel (mesmo padrao ja usado para "À COMEÇAR").
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")
lines = text.split("\r\n")


def find_line_exact(s, start=0):
    for i in range(start, len(lines)):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


# ---------------------------------------------------------------
# 1) cmdNovo_Click - resetar estado dos botoes de servico
# ---------------------------------------------------------------
i = find_line_exact("cmdApagar.Enabled = False")
lines[i] = (
    lines[i]
    + "\r\n"
    + "cmdAdicionarServicosAuto.Enabled = True\r\n"
    + "cmdRemoverServicosAuto.Enabled = True\r\n"
    + "cmdEditarServicosAuto.Enabled = False\r\n"
    + 'vCodItemServicoEditando = ""\r\n'
    + 'cboMecanicoServ.Text = ""\r\n'
    + 'vCodMecanicoServ = ""'
)

# ---------------------------------------------------------------
# 2) Grid_Servicos_DblClick - preencher cboServicosAuto tambem
# ---------------------------------------------------------------
i = find_line_exact("vServico = Grid_Servicos.TextMatrix(Grid_Servicos.Row, 3)")
lines[i] = lines[i] + "\r\ncboServicosAuto.Text = vServico"

# ---------------------------------------------------------------
# 3) cmdEditarOS_Click - stProdSer invisivel quando status = EM EXECUÇÃO
# ---------------------------------------------------------------
start_eos, end_eos = find_sub("cmdEditarOS_Click")

marker = None
for idx in range(start_eos, end_eos):
    if lines[idx].strip().startswith("If (Trim(Grid_OS.TextMatrix(posit, 1)))"):
        marker = idx
        break
assert marker is not None, "bloco stProdSer nao encontrado em cmdEditarOS_Click"

# encontra o "Else" que fecha esse If (nao o End If interno do vTipoOS)
inner_endif = None
for idx in range(marker + 1, end_eos):
    if lines[idx].strip() == "End If":
        inner_endif = idx
        break
assert inner_endif is not None

outer_else = inner_endif + 1
assert lines[outer_else].strip() == "Else", lines[outer_else]

new_lines = [
    'ElseIf (Trim(Grid_OS.TextMatrix(posit, 1))) = ("EM EXECUÇÃO") Then',
    "    stProdSer.Visible = False",
]
lines[outer_else:outer_else] = new_lines

# ---------------------------------------------------------------
# Grava
# ---------------------------------------------------------------
out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
print("bytes originais:", len(raw), "bytes finais:", len(out))
