# -*- coding: utf-8 -*-
"""
Vendas_Consulta_Geral_Parcelas.frm - FormatarGrid_Parcelas:
- Corrige bug pre-existente: o loop "MUDAR COR DE FONTE DA COLUNA"
  mexia na coluna 5 (DESC.) em vez da 9 (STATUS) - por isso DESC.
  sempre aparecia vermelho/negrito (a condicao "TextMatrix(i,5)=PAGO"
  nunca era verdadeira, ja que a coluna 5 guarda um numero formatado).
- PGTO (8) e STATUS (9): fundo cinza claro.
- STATUS = PAGO -> fonte verde escura negrito; A PAGAR -> vermelho negrito.
- PGTO: mesma formatacao de STATUS quando PAGO; preto sem negrito quando
  A PAGAR.
- DESC. (5): fonte preta, sem negrito (tira o bug antigo).
- VENC. (2): negrito, vermelho escuro.
"""

PATH = r"C:\projeto\Compartilhado\Forms\Vendas_Consulta_Geral_Parcelas.frm"

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


i = find_line_exact("      'MUDAR COR DE FONTE DA COLUNA")
end = find_line_exact("      Next", i)
old = lines[i : end + 1]
expected = [
    "      'MUDAR COR DE FONTE DA COLUNA",
    "      For i = 1 To .Rows - 1",
    "         .Row = i",
    "         .Col = 5",
    '         If .TextMatrix(i, 5) = "PAGO" Then',
    "            .CellForeColor = vbBlue",
    "         Else",
    "            .CellForeColor = vbRed",
    "         End If",
    "         ",
    "         .CellFontBold = True",
    "      Next",
]
assert old == expected, old

novo = [
    "      'fundo cinza claro nas colunas PGTO e STATUS",
    "      For i = 0 To .Rows - 1",
    "         .Row = i",
    "         .Col = 8",
    "         .CellBackColor = &HE0E0E0",
    "         .Col = 9",
    "         .CellBackColor = &HE0E0E0",
    "      Next",
    "      ",
    "      'STATUS: PAGO = verde escuro negrito / À PAGAR = vermelho negrito",
    "      'PGTO: mesma formatação de STATUS quando PAGO; preto sem negrito quando À PAGAR",
    "      For i = 1 To .Rows - 1",
    "         .Row = i",
    "         .Col = 9",
    '         If .TextMatrix(i, 9) = "PAGO" Then',
    "            .CellForeColor = RGB(0, 100, 0)",
    "         Else",
    "            .CellForeColor = vbRed",
    "         End If",
    "         .CellFontBold = True",
    "         ",
    "         .Col = 8",
    '         If .TextMatrix(i, 9) = "PAGO" Then',
    "            .CellForeColor = RGB(0, 100, 0)",
    "            .CellFontBold = True",
    "         Else",
    "            .CellForeColor = vbBlack",
    "            .CellFontBold = False",
    "         End If",
    "      Next",
    "      ",
    "      'DESC.: fonte preta, sem negrito",
    "      For i = 1 To .Rows - 1",
    "         .Row = i",
    "         .Col = 5",
    "         .CellForeColor = vbBlack",
    "         .CellFontBold = False",
    "      Next",
    "      ",
    "      'VENC.: negrito, vermelho escuro",
    "      For i = 1 To .Rows - 1",
    "         .Row = i",
    "         .Col = 2",
    "         .CellForeColor = RGB(139, 0, 0)",
    "         .CellFontBold = True",
    "      Next",
]
lines[i : end + 1] = novo

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
