# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - MostrarGrid_OS_Refinado: adiciona os mesmos filtros
SITUACAO (cboConsultaStatus) e var_STATUS (cboConsultaMostrar) que ja
existem em MostrarGrid_OS, injetando no WHERE antes de varTipoPagamento.
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


i = find_line_exact("Private Sub MostrarGrid_OS_Refinado()")
end = find_line_exact("End Sub", i)

j = find_line_exact("'forma de pagamento (TODOS/À VISTA/À PRAZO)", i, end)

novo_calc = [
    "Dim SITUACAO As String",
    "Dim var_STATUS As String",
    "",
    "'Status",
    'If cboConsultaStatus.Text = "TODOS" Then',
    '   SITUACAO = ""',
    'ElseIf cboConsultaStatus.Text = "À COMEÇAR" Then',
    "   SITUACAO = \"AND (os.status = 'À COMEÇAR') \"",
    'ElseIf cboConsultaStatus.Text = "EM EXECUÇÃO" Then',
    "   SITUACAO = \"AND (os.status = 'EM EXECUÇÃO') \"",
    'ElseIf cboConsultaStatus.Text = "AGUARDANDO" Then',
    "   SITUACAO = \"AND (os.status = 'AGUARDANDO') \"",
    'ElseIf cboConsultaStatus.Text = "TERMINADO" Then',
    "   SITUACAO = \"AND (os.status = 'TERMINADO') \"",
    "End If",
    "",
    "'Situação",
    'If cboConsultaMostrar.Text = "TODOS" Then',
    '   var_STATUS = ""',
    'ElseIf cboConsultaMostrar.Text = "ABERTOS" Then',
    '   var_STATUS = "AND (status_os = 0) "',
    'ElseIf cboConsultaMostrar.Text = "FECHADOS" Then',
    '   var_STATUS = "AND (status_os = 1) "',
    "End If",
    "",
]
lines[j:j] = novo_calc
end = find_line_exact("End Sub", i)

m = find_line_exact('   "WHERE (" & campoBusca & " = \'" & txtFiltroRefinado.Text & "\') " & _', i, end)
n = find_line_exact('   varTipoPagamento & "ORDER BY os.DATA_TERMINO DESC"', m, m + 3)
lines[n] = "   SITUACAO & var_STATUS & varTipoPagamento & \"ORDER BY os.DATA_TERMINO DESC\""

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
