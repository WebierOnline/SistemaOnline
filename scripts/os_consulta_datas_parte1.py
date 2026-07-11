# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - parte 1:
A) renomeia chameleonButton1 -> cmdCal2
B) mskDataConsulta e mskPeriodoFim: mascara "##/##/####" -> "##/##/##"
C) AtualizarCamposCriterios: atualiza lblCriterio(5).Caption e mostra/
   esconde optDataEntrada/optDataTermino junto com os campos de data
D) cmdCal1_Click / cmdCal2_Click (mesmo padrao do OS_Recapadora.cmdCal1)
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
# A) renomear chameleonButton1 -> cmdCal2
# ---------------------------------------------------------------
i = find_line_exact("      Begin ChamaleonBtn.chameleonButton chameleonButton1 ")
lines[i] = "      Begin ChamaleonBtn.chameleonButton cmdCal2 "

# ---------------------------------------------------------------
# B) mascara de 2 digitos no ano
# ---------------------------------------------------------------
i = find_line_exact("      Begin MSMask.MaskEdBox mskDataConsulta ")
j = find_line_exact('         Mask            =   "##/##/####"', i, i + 15)
lines[j] = '         Mask            =   "##/##/##"'

i = find_line_exact("      Begin MSMask.MaskEdBox mskPeriodoFim ")
j = find_line_exact('         Mask            =   "##/##/####"', i, i + 15)
lines[j] = '         Mask            =   "##/##/##"'

# ---------------------------------------------------------------
# C) AtualizarCamposCriterios - reescreve por completo
# ---------------------------------------------------------------
i = find_line_exact("Private Sub AtualizarCamposCriterios()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub AtualizarCamposCriterios()",
    "cboLocalizar.Visible = False",
    "mskDataConsulta.Visible = False",
    "mskPeriodoInicio.Visible = False",
    "lblPeriodoAte.Visible = False",
    "mskPeriodoFim.Visible = False",
    "cboMesConsulta.Visible = False",
    "cboAnoConsulta.Visible = False",
    "",
    'If cboConsultaCriterios.Text = "TODOS" Then',
    '   cboLocalizar.Text = ""',
    'ElseIf cboConsultaCriterios.Text = "CLIENTE" Or cboConsultaCriterios.Text = "CÓD. OS" Then',
    "   cboLocalizar.Visible = True",
    'ElseIf cboConsultaCriterios.Text = "DATA" Then',
    "   mskDataConsulta.Visible = True",
    'ElseIf cboConsultaCriterios.Text = "PERÍODO" Then',
    "   mskPeriodoInicio.Visible = True",
    "   lblPeriodoAte.Visible = True",
    "   mskPeriodoFim.Visible = True",
    'ElseIf cboConsultaCriterios.Text = "MENSAL" Then',
    "   cboMesConsulta.Visible = True",
    "   cboAnoConsulta.Visible = True",
    "End If",
    "End Sub",
]
assert old == expected, old

novo = [
    "Private Sub AtualizarCamposCriterios()",
    "cboLocalizar.Visible = False",
    "mskDataConsulta.Visible = False",
    "cmdCal1.Visible = False",
    "mskPeriodoInicio.Visible = False",
    "lblPeriodoAte.Visible = False",
    "mskPeriodoFim.Visible = False",
    "cmdCal2.Visible = False",
    "cboMesConsulta.Visible = False",
    "cboAnoConsulta.Visible = False",
    "optDataEntrada.Visible = False",
    "optDataTermino.Visible = False",
    "",
    'If cboConsultaCriterios.Text = "TODOS" Then',
    '   cboLocalizar.Text = ""',
    '   lblCriterio(5).Caption = ""',
    'ElseIf cboConsultaCriterios.Text = "CLIENTE" Then',
    "   cboLocalizar.Visible = True",
    '   lblCriterio(5).Caption = "Cliente:"',
    'ElseIf cboConsultaCriterios.Text = "CÓD. OS" Then',
    "   cboLocalizar.Visible = True",
    '   lblCriterio(5).Caption = "Código:"',
    'ElseIf cboConsultaCriterios.Text = "DATA" Then',
    "   mskDataConsulta.Visible = True",
    "   cmdCal1.Visible = True",
    "   optDataEntrada.Visible = True",
    "   optDataTermino.Visible = True",
    '   lblCriterio(5).Caption = "Data:"',
    'ElseIf cboConsultaCriterios.Text = "PERÍODO" Then',
    "   mskPeriodoInicio.Visible = True",
    "   lblPeriodoAte.Visible = True",
    "   mskPeriodoFim.Visible = True",
    "   cmdCal2.Visible = True",
    "   optDataEntrada.Visible = True",
    "   optDataTermino.Visible = True",
    '   lblCriterio(5).Caption = "Período:"',
    'ElseIf cboConsultaCriterios.Text = "MENSAL" Then',
    "   cboMesConsulta.Visible = True",
    "   cboAnoConsulta.Visible = True",
    "   optDataEntrada.Visible = True",
    "   optDataTermino.Visible = True",
    '   lblCriterio(5).Caption = "Mês/Ano:"',
    "End If",
    "End Sub",
]
lines[i : end + 1] = novo

# ---------------------------------------------------------------
# D) cmdCal1_Click / cmdCal2_Click - inseridos apos cmdFechar_Click
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cmdFechar_Click()")
end = find_line_exact("End Sub", i)

novos_cal = [
    "",
    "Private Sub cmdCal1_Click()",
    "Dim varData As Variant",
    "Dim fCal As Calendario",
    "",
    "varData = Empty",
    "",
    "Set fCal = New Calendario",
    "fCal.Show vbModal",
    "",
    "varData = fCal.DateSelected",
    "",
    "Unload fCal",
    "Set fCal = Nothing",
    "",
    "If Not IsDate(varData) Then Exit Sub",
    "If varData = 0 Then Exit Sub",
    "",
    'mskDataConsulta = Format(varData, "dd/mm/yy")',
    "End Sub",
    "",
    "Private Sub cmdCal2_Click()",
    "Dim varData As Variant",
    "Dim fCal As Calendario",
    "",
    "varData = Empty",
    "",
    "Set fCal = New Calendario",
    "fCal.Show vbModal",
    "",
    "varData = fCal.DateSelected",
    "",
    "Unload fCal",
    "Set fCal = Nothing",
    "",
    "If Not IsDate(varData) Then Exit Sub",
    "If varData = 0 Then Exit Sub",
    "",
    'mskPeriodoFim = Format(varData, "dd/mm/yy")',
    "End Sub",
]
lines[end + 1 : end + 1] = novos_cal

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 1 aplicada")
