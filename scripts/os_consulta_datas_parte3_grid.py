# -*- coding: utf-8 -*-
"""
OS_Consulta.frm - parte 3: adiciona colunas ENTRADA e TERMINO no Grid,
logo apos COD., empurrando as demais colunas +2 posicoes.
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


i = find_line_exact("Private Sub FormatarGrid_OS(rTabela As ADODB.Recordset)")
end = find_line_exact("End Sub", i)
sub_lines = lines[i : end + 1]

# ---------------------------------------------------------------
# FormatString + ColWidth
# ---------------------------------------------------------------
old_header = [
    '   .FormatString = "^CÓD.|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL"',
    "   .ColWidth(0) = 650",
    "   .ColWidth(1) = 1250",
    "   .ColWidth(2) = 1000",
    "   .ColWidth(3) = 5350",
    "   .ColWidth(4) = 750",
    "   .ColWidth(5) = 750",
    "   .ColWidth(6) = 850",
    "   .ColWidth(7) = 650",
    "   .ColWidth(8) = 850",
]
new_header = [
    '   .FormatString = "^CÓD.|^ENTRADA|^TERMINO|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL"',
    "   .ColWidth(0) = 650",
    "   .ColWidth(1) = 1100",
    "   .ColWidth(2) = 1100",
    "   .ColWidth(3) = 1250",
    "   .ColWidth(4) = 1000",
    "   .ColWidth(5) = 5350",
    "   .ColWidth(6) = 750",
    "   .ColWidth(7) = 750",
    "   .ColWidth(8) = 850",
    "   .ColWidth(9) = 650",
    "   .ColWidth(10) = 850",
]
idx = sub_lines.index(old_header[0])
assert sub_lines[idx : idx + len(old_header)] == old_header
sub_lines[idx : idx + len(old_header)] = new_header

# ---------------------------------------------------------------
# ColAlignment (dentro do Do While) - shift +2
# ---------------------------------------------------------------
old_align = [
    "         .ColAlignment(3) = 1",
    "         .ColAlignment(6) = 0",
    "         .ColAlignment(5) = 0",
    "         .ColAlignment(6) = 6",
    "         .ColAlignment(7) = 6",
    "         .ColAlignment(8) = 6",
]
new_align = [
    "         .ColAlignment(5) = 1",
    "         .ColAlignment(8) = 0",
    "         .ColAlignment(7) = 0",
    "         .ColAlignment(8) = 6",
    "         .ColAlignment(9) = 6",
    "         .ColAlignment(10) = 6",
]
idx = sub_lines.index(old_align[0])
assert sub_lines[idx : idx + len(old_align)] == old_align
sub_lines[idx : idx + len(old_align)] = new_align

# ---------------------------------------------------------------
# TextMatrix - insere ENTRADA/TERMINO (novas col 1 e 2), desloca o
# resto (+2), usando os aliases DATA_ENTRADA/DATA_TERMINO do SELECT
# ---------------------------------------------------------------
old_matrix = [
    '         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "0000")',
    '         .TextMatrix(.Rows - 1, 1) = rTabela("var_status")',
    '         .TextMatrix(.Rows - 1, 2) = rTabela("var_status_os") & ""',
    "         ",
    '         If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then',
    '            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo")) & " / " & ValidateNull(rTabela("ano"))',
    '         ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then',
    '            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))',
    '         ElseIf vTipoOS = "Comunicação Visual" Then',
    '            .TextMatrix(.Rows - 1, 3) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))',
    "         End If",
    '         .TextMatrix(.Rows - 1, 4) = ValidateNull(rTabela("TIPO_PAGAMENTO"))',
    '         .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("PAGAMENTO"))',
    '         .TextMatrix(.Rows - 1, 6) = Format(rTabela("SUBTOTAL"), ocMONEY)',
    '         .TextMatrix(.Rows - 1, 7) = Format(rTabela("ValorDescReal"), ocMONEY)',
    '         .TextMatrix(.Rows - 1, 8) = Format(rTabela("TOTAL"), ocMONEY)',
]
new_matrix = [
    '         .TextMatrix(.Rows - 1, 0) = Format(rTabela("cod_os"), "0000")',
    '         If IsNull(rTabela("DATA_ENTRADA")) Then',
    "            .TextMatrix(.Rows - 1, 1) = \"\"",
    "         Else",
    '            .TextMatrix(.Rows - 1, 1) = Format(rTabela("DATA_ENTRADA"), "dd/mm/yy")',
    "         End If",
    '         If IsNull(rTabela("DATA_TERMINO")) Then',
    "            .TextMatrix(.Rows - 1, 2) = \"\"",
    "         Else",
    '            .TextMatrix(.Rows - 1, 2) = Format(rTabela("DATA_TERMINO"), "dd/mm/yy")',
    "         End If",
    '         .TextMatrix(.Rows - 1, 3) = rTabela("var_status")',
    '         .TextMatrix(.Rows - 1, 4) = rTabela("var_status_os") & ""',
    "         ",
    '         If vTipoOS = "Automóveis" Or vTipoOS = "Motocicletas" Or vTipoOS = "Recapadora" Then',
    '            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo")) & " / " & ValidateNull(rTabela("ano"))',
    '         ElseIf vTipoOS = "Informática" Or vTipoOS = "Celular" Then',
    '            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))',
    '         ElseIf vTipoOS = "Comunicação Visual" Then',
    '            .TextMatrix(.Rows - 1, 5) = ValidateNull(rTabela("nome")) & " / " & ValidateNull(rTabela("equipamento")) & " / " & ValidateNull(rTabela("fabricante")) & " / " & ValidateNull(rTabela("modelo"))',
    "         End If",
    '         .TextMatrix(.Rows - 1, 6) = ValidateNull(rTabela("TIPO_PAGAMENTO"))',
    '         .TextMatrix(.Rows - 1, 7) = ValidateNull(rTabela("PAGAMENTO"))',
    '         .TextMatrix(.Rows - 1, 8) = Format(rTabela("SUBTOTAL"), ocMONEY)',
    '         .TextMatrix(.Rows - 1, 9) = Format(rTabela("ValorDescReal"), ocMONEY)',
    '         .TextMatrix(.Rows - 1, 10) = Format(rTabela("TOTAL"), ocMONEY)',
]
idx = sub_lines.index(old_matrix[0])
assert sub_lines[idx : idx + len(old_matrix)] == old_matrix, sub_lines[idx : idx + len(old_matrix)]
sub_lines[idx : idx + len(old_matrix)] = new_matrix

# ---------------------------------------------------------------
# cor da coluna 'ABERTO/FECHADO' e status - eram colunas 2 e 1, agora 4 e 3
# ---------------------------------------------------------------
old_cor1 = [
    "   For i = 1 To .Rows - 1",
    '      If UCase(Trim(.TextMatrix(i, 2))) = UCase("ABERTO") Then',
    "         aCor = vbBlue",
    "      Else",
    "         aCor = vbRed",
    "      End If",
    "      ",
    "      .Col = 2 'a coluna do aberto ou fechado",
    "      .Row = i",
    "      .CellForeColor = aCor",
    "   Next",
]
new_cor1 = [
    "   For i = 1 To .Rows - 1",
    '      If UCase(Trim(.TextMatrix(i, 4))) = UCase("ABERTO") Then',
    "         aCor = vbBlue",
    "      Else",
    "         aCor = vbRed",
    "      End If",
    "      ",
    "      .Col = 4 'a coluna do aberto ou fechado",
    "      .Row = i",
    "      .CellForeColor = aCor",
    "   Next",
]
idx = sub_lines.index(old_cor1[0])
assert sub_lines[idx : idx + len(old_cor1)] == old_cor1, sub_lines[idx : idx + len(old_cor1)]
sub_lines[idx : idx + len(old_cor1)] = new_cor1
idx_apos_cor1 = idx + len(new_cor1)

old_cor2 = [
    "   For i = 1 To .Rows - 1",
    '      If UCase(Trim(.TextMatrix(i, 1))) = UCase("À COMEÇAR") Then',
    "         aCor = vbBlack",
    '      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("EM EXECUÇÃO") Then',
    "         aCor = vbGreen",
    '      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("AGUARDANDO") Then',
    "         aCor = vbBlue",
    '      ElseIf UCase(Trim(.TextMatrix(i, 1))) = UCase("TERMINADO") Then',
    "         aCor = vbRed",
    "      End If",
    "      ",
    "      .Col = 1 'a coluna do aberto ou fechado",
    "      .Row = i",
    "      .CellForeColor = aCor",
    "   Next",
]
new_cor2 = [
    "   For i = 1 To .Rows - 1",
    '      If UCase(Trim(.TextMatrix(i, 3))) = UCase("À COMEÇAR") Then',
    "         aCor = vbBlack",
    '      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("EM EXECUÇÃO") Then',
    "         aCor = vbGreen",
    '      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("AGUARDANDO") Then',
    "         aCor = vbBlue",
    '      ElseIf UCase(Trim(.TextMatrix(i, 3))) = UCase("TERMINADO") Then',
    "         aCor = vbRed",
    "      End If",
    "      ",
    "      .Col = 3 'a coluna do aberto ou fechado",
    "      .Row = i",
    "      .CellForeColor = aCor",
    "   Next",
]
idx = None
for k in range(idx_apos_cor1, len(sub_lines)):
    if sub_lines[k] == old_cor2[0]:
        idx = k
        break
assert idx is not None, "old_cor2 nao encontrado apos cor1"
assert sub_lines[idx : idx + len(old_cor2)] == old_cor2, sub_lines[idx : idx + len(old_cor2)]
sub_lines[idx : idx + len(old_cor2)] = new_cor2

# ---------------------------------------------------------------
# lblTotalConsulta usa SomaGrid(Grid, 8) - coluna TOTAL agora e 10
# ---------------------------------------------------------------
old_soma = "lblTotalConsulta.Caption = Format(SomaGrid(Grid, 8), ocMONEY)"
new_soma = "lblTotalConsulta.Caption = Format(SomaGrid(Grid, 10), ocMONEY)"
idx = sub_lines.index(old_soma)
sub_lines[idx] = new_soma

lines[i : end + 1] = sub_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - parte 3 (grid ENTRADA/TERMINO) aplicada")
