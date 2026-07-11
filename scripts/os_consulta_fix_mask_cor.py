# -*- coding: utf-8 -*-
"""
OS_Consulta.frm:
1) mskPeriodoInicio: mascara "##/##/####" -> "##/##/##" (2 digitos de
   ano) - agora que cmdCal1 tambem preenche esse campo no formato
   "dd/mm/yy", precisa bater com o formato do cmdCal2/mskPeriodoFim.
2) Colunas TERMINO (2) e FINANC. (4) do Grid ganham fundo cinza claro
   (&HE0E0E0, mesmo tom ja usado em outros grids do projeto).
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
# 1) mascara de mskPeriodoInicio
# ---------------------------------------------------------------
i = find_line_exact("      Begin MSMask.MaskEdBox mskPeriodoInicio ")
j = find_line_exact('         Mask            =   "##/##/####"', i, i + 15)
lines[j] = '         Mask            =   "##/##/##"'

# ---------------------------------------------------------------
# 2) fundo cinza claro nas colunas TERMINO (2) e FINANC. (4)
# ---------------------------------------------------------------
i = find_line_exact("Private Sub FormatarGrid_OS(rTabela As ADODB.Recordset)")
end = find_line_exact("End Sub", i)
anchor = "   .Redraw = True"
j = find_line_exact(anchor, i, end)
novo_bloco = """   'colunas TERMINO e FINANC. com fundo cinza claro
   For i = 0 To .Rows - 1
      .Row = i
      .Col = 2
      .CellBackColor = &HE0E0E0
      .Col = 4
      .CellBackColor = &HE0E0E0
   Next

""" + anchor
lines[j] = novo_bloco

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
