# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - filtro refinado por Placa/Chassi:
1) optFiltroSimples/optFiltroRefinado alternam Enabled de
   frmConsultaSimples/frmConsultaRefina (mutuamente exclusivos); estado
   inicial (Filtro Simples marcado) setado em Form_Load.
2) chkPlaca/chkChassi mutuamente exclusivos.
3) cmdExibir_Click passa a chamar MostrarGrid_OS_Refinado (nova sub)
   quando o Filtro Refinado estiver ativo, consultando OS_Equipamento_Auto
   por PLACA ou CHASSI (so faz sentido p/ veiculos), ordenado por
   DATA_TERMINO.
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
# 1) Form_Load: estado inicial do toggle simples/refinado
# ---------------------------------------------------------------
i = find_line_exact("Private Sub Form_Load()")
j = find_line_exact("Set moCombo = New cComboHelper", i, i + 5)
lines[j] = (
    "Set moCombo = New cComboHelper\r\n"
    "frmConsultaRefina.Enabled = False"
)

# ---------------------------------------------------------------
# 2) cmdExibir_Click: chama a rotina certa conforme o filtro ativo
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cmdExibir_Click()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
assert old == ["Private Sub cmdExibir_Click()", "MostrarGrid_OS", "End Sub"], old
novo = [
    "Private Sub cmdExibir_Click()",
    "If optFiltroRefinado.Value = True Then",
    "   MostrarGrid_OS_Refinado",
    "Else",
    "   MostrarGrid_OS",
    "End If",
    "End Sub",
]
lines[i : end + 1] = novo

# ---------------------------------------------------------------
# 3) novos handlers - inseridos logo apos cmdExibir_Click
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cmdExibir_Click()")
end = find_line_exact("End Sub", i)

novos_subs = [
    "",
    "Private Sub optFiltroSimples_Click()",
    "frmConsultaSimples.Enabled = True",
    "frmConsultaRefina.Enabled = False",
    "End Sub",
    "",
    "Private Sub optFiltroRefinado_Click()",
    "frmConsultaSimples.Enabled = False",
    "frmConsultaRefina.Enabled = True",
    "End Sub",
    "",
    "Private Sub chkPlaca_Click()",
    "If chkPlaca.Value = 1 Then chkChassi.Value = 0",
    "End Sub",
    "",
    "Private Sub chkChassi_Click()",
    "If chkChassi.Value = 1 Then chkPlaca.Value = 0",
    "End Sub",
    "",
    "Private Sub MostrarGrid_OS_Refinado()",
    "Dim totalRegistros As Long",
    "Dim campoBusca As String",
    "",
    'If vTipoOS <> "Automóveis" And vTipoOS <> "Motocicletas" And vTipoOS <> "Recapadora" Then',
    '   MsgBox "Consulta por Placa/Chassi disponível apenas para veículos!", vbInformation, "Aviso do Sistema"',
    "   Exit Sub",
    "End If",
    "",
    "If chkPlaca.Value = 1 Then",
    '   campoBusca = "OS_Equipamento_Auto.PLACA"',
    "ElseIf chkChassi.Value = 1 Then",
    '   campoBusca = "OS_Equipamento_Auto.CHASSI"',
    "Else",
    '   MsgBox "Selecione Placa ou Chassi!", vbInformation, "Aviso do Sistema"',
    "   Exit Sub",
    "End If",
    "",
    'If txtFiltroRefinado.Text = "" Then',
    '   MsgBox "Digite a Placa ou o Chassi para consultar!", vbInformation, "Aviso do Sistema"',
    "   Exit Sub",
    "End If",
    "",
    "sSQL = \"SELECT DISTINCT OS.COD_OS, cliente.Nome, os.DATA_ENTRADA, os.DATA_TERMINO, OS_Equipamento_Auto.fabricante, OS_Equipamento_Auto.ano, OS_Equipamento_Auto.modelo, os.status AS var_status, CASE status_os WHEN 1 THEN 'FECHADO' WHEN 0 THEN 'ABERTO' END AS var_status_os, os.TIPO_PAGAMENTO, os.PAGAMENTO, os.SUBTOTAL, os.ValorDescReal, os.TOTAL \" & _",
    '   "FROM cliente INNER JOIN OS ON cliente.CODIGO = OS.COD_CLIENTE INNER JOIN OS_Equipamento_Auto ON OS.COD_OS = OS_Equipamento_Auto.COD_OS " & _',
    '   "WHERE (" & campoBusca & " = \'" & txtFiltroRefinado.Text & "\') " & _',
    '   "ORDER BY os.DATA_TERMINO DESC"',
    "",
    "Set r = dbData.OpenRecordset(sSQL, totalRegistros)",
    "",
    "FormatarGrid_OS r",
    "",
    "printSQL = sSQL",
    "",
    'lblQuant.Caption = "QUANTIDADE: " & Format(totalRegistros, "000")',
    "",
    "If r.State <> 0 Then r.Close",
    "Set r = Nothing",
    "End Sub",
]
lines[end + 1 : end + 1] = novos_subs

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
