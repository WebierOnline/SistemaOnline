# -*- coding: utf-8 -*-
"""
Parte 2: reescreve cboConsultaCriterios_Click/_Validate com um helper
AtualizarCamposCriterios, e adiciona cboMesConsulta_GotFocus /
cboAnoConsulta_GotFocus.
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
# cboConsultaCriterios_Click + _Validate -> reescritos, com helper novo
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cboConsultaCriterios_Click()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub cboConsultaCriterios_Click()",
    'If cboConsultaCriterios.Text = "TODOS" Then',
    '   cboLocalizar.Text = ""',
    "   cboLocalizar.Visible = False",
    "   MostrarGrid_OS",
    "Else",
    "   cboLocalizar.Visible = True",
    "   cboLocalizar.SetFocus",
    "End If",
    "End Sub",
]
assert old == expected, old

i2 = find_line_exact("Private Sub cboConsultaCriterios_Validate(Cancel As Boolean)")
end2 = find_line_exact("End Sub", i2)
old2 = lines[i2 : end2 + 1]
expected2 = [
    "Private Sub cboConsultaCriterios_Validate(Cancel As Boolean)",
    'If cboConsultaCriterios.Text = "TODOS" Then',
    '   cboLocalizar.Text = ""',
    "   cboLocalizar.Visible = False",
    "Else",
    "   cboLocalizar.Visible = True",
    "End If",
    "End Sub",
]
assert old2 == expected2, old2

# substitui os dois juntos (end2 depois de end, entao processa de tras pra frente)
novo2 = [
    "Private Sub cboConsultaCriterios_Validate(Cancel As Boolean)",
    "AtualizarCamposCriterios",
    "End Sub",
]
lines[i2 : end2 + 1] = novo2

novo1 = [
    "Private Sub cboConsultaCriterios_Click()",
    "AtualizarCamposCriterios",
    'If cboConsultaCriterios.Text = "TODOS" Then',
    "   MostrarGrid_OS",
    'ElseIf cboConsultaCriterios.Text = "CLIENTE" Or cboConsultaCriterios.Text = "CÓD. OS" Then',
    "   cboLocalizar.SetFocus",
    'ElseIf cboConsultaCriterios.Text = "DATA" Then',
    "   mskDataConsulta.SetFocus",
    'ElseIf cboConsultaCriterios.Text = "PERÍODO" Then',
    "   mskPeriodoInicio.SetFocus",
    'ElseIf cboConsultaCriterios.Text = "MENSAL" Then',
    "   cboMesConsulta.SetFocus",
    "End If",
    "End Sub",
    "",
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
lines[i : end + 1] = novo1

# ---------------------------------------------------------------
# cboMesConsulta_GotFocus / cboAnoConsulta_GotFocus -> inseridos apos
# cboLocalizar_LostFocus (fim do bloco dos combos de criterio)
# ---------------------------------------------------------------
i3 = find_line_exact("Private Sub cboLocalizar_LostFocus()")
end3 = find_line_exact("End Sub", i3)

novos_subs = [
    "",
    "Private Sub cboMesConsulta_GotFocus()",
    "cboMesConsulta.Clear",
    'cboMesConsulta.AddItem "Janeiro"',
    'cboMesConsulta.AddItem "Fevereiro"',
    'cboMesConsulta.AddItem "Março"',
    'cboMesConsulta.AddItem "Abril"',
    'cboMesConsulta.AddItem "Maio"',
    'cboMesConsulta.AddItem "Junho"',
    'cboMesConsulta.AddItem "Julho"',
    'cboMesConsulta.AddItem "Agosto"',
    'cboMesConsulta.AddItem "Setembro"',
    'cboMesConsulta.AddItem "Outubro"',
    'cboMesConsulta.AddItem "Novembro"',
    'cboMesConsulta.AddItem "Dezembro"',
    "moCombo.AttachTo cboMesConsulta",
    "End Sub",
    "",
    "Private Sub cboAnoConsulta_GotFocus()",
    "Dim iAno As Integer",
    "Dim i As Integer",
    "cboAnoConsulta.Clear",
    "iAno = Year(Date)",
    "For i = iAno - 5 To iAno + 1",
    "   cboAnoConsulta.AddItem i",
    "Next",
    "moCombo.AttachTo cboAnoConsulta",
    "End Sub",
]
lines[end3 + 1 : end3 + 1] = novos_subs

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 2 aplicada")
