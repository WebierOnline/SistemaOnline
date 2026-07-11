# -*- coding: utf-8 -*-
"""
OS_Consulta.frm:
1) cmdCal1 tambem aparece (e passa a preencher mskPeriodoInicio) quando
   o criterio for PERIODO - ele e o mesmo botao que preenche
   mskDataConsulta no criterio DATA (os dois campos ocupam a mesma
   posicao na tela, so um fica visivel por vez).
2) Form_Load: criterio padrao passa a ser MENSAL, com mes/ano atuais
   ja carregados em cboMesConsulta/cboAnoConsulta.
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
# 1a) AtualizarCamposCriterios: cmdCal1 tambem visivel no PERIODO
# ---------------------------------------------------------------
i_atualizar = find_line_exact("Private Sub AtualizarCamposCriterios()")
end_atualizar = find_line_exact("End Sub", i_atualizar)
i = find_line_exact('ElseIf cboConsultaCriterios.Text = "PERÍODO" Then', i_atualizar, end_atualizar)
j = find_line_exact("   mskPeriodoInicio.Visible = True", i, i + 5)
lines[j] = "   mskPeriodoInicio.Visible = True\r\n   cmdCal1.Visible = True"

# ---------------------------------------------------------------
# 1b) cmdCal1_Click: contextual - PERIODO preenche mskPeriodoInicio,
#     demais (DATA) preenchem mskDataConsulta
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cmdCal1_Click()")
end = find_line_exact("End Sub", i)
old_last = "mskDataConsulta = Format(varData, \"dd/mm/yy\")"
j = find_line_exact(old_last, i, end)
lines[j] = (
    'If cboConsultaCriterios.Text = "PERÍODO" Then\r\n'
    '   mskPeriodoInicio = Format(varData, "dd/mm/yy")\r\n'
    "Else\r\n"
    '   mskDataConsulta = Format(varData, "dd/mm/yy")\r\n'
    "End If"
)

# ---------------------------------------------------------------
# 2) Form_Load: MENSAL como criterio padrao + mes/ano atuais
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Form_Load()")
end = find_line_exact("End Sub", i)
old_block = [
    "cboConsultaMostrar.ListIndex = 0",
    "cboConsultaStatus.ListIndex = 0",
    "cboConsultaCriterios.ListIndex = 0",
    "AtualizarCamposCriterios",
    "cboTipoServico.ListIndex = 0",
    "cboIndice.ListIndex = 0",
]
j = find_line_exact(old_block[0], i, end)
actual = lines[j : j + len(old_block)]
assert actual == old_block, actual

novo_block = [
    "cboConsultaMostrar.ListIndex = 0",
    "cboConsultaStatus.ListIndex = 0",
    "cboTipoServico.ListIndex = 0",
    "cboIndice.ListIndex = 0",
    "",
    "cboMesConsulta_GotFocus",
    "cboMesConsulta.ListIndex = Month(Date) - 1",
    "cboAnoConsulta_GotFocus",
    "cboAnoConsulta.Text = Year(Date)",
    "",
    'cboConsultaCriterios.Text = "MENSAL"',
    "AtualizarCamposCriterios",
]
lines[j : j + len(old_block)] = novo_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
