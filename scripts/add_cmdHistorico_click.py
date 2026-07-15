# -*- coding: utf-8 -*-
"""
frmBuscarPlaca.frm: adiciona cmdHistorico_Click - abre OS_Consulta ja
filtrado (Refinado/Placa) pela placa selecionada no grid.
"""

PATH = r"C:\projeto\OrdemServico\Forms\frmBuscarPlaca.frm"

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


anchor = find_line_exact("Private Sub cmdFechar_Click()")

new_lines = [
    "Private Sub cmdHistorico_Click()",
    "    If lstVeiculos.Row < 1 Then",
    '        MsgBox "Selecione um ve\xedculo.", vbInformation, "Aviso do Sistema"',
    "        Exit Sub",
    "    End If",
    "",
    "    Dim sPlaca As String",
    "    sPlaca = Trim(lstVeiculos.TextMatrix(lstVeiculos.Row, 5))",
    '    If sPlaca = "" Then',
    '        MsgBox "Ve\xedculo sem placa cadastrada.", vbInformation, "Aviso do Sistema"',
    "        Exit Sub",
    "    End If",
    "",
    "    OS_Consulta.sPlacaBusca = sPlaca",
    "    OS_Consulta.Show vbModal",
    "    Unload OS_Consulta",
    "End Sub",
    "",
]

lines[anchor:anchor] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - cmdHistorico_Click inserido antes da linha {anchor}")
