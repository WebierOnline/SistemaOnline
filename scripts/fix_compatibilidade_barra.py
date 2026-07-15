# -*- coding: utf-8 -*-
"""
OS_Consulta_Pecas.frm: a coluna de compatibilidade (chkCompartibilidade)
exibe "/ CORSA (98/03), / CLASSIC(2004/15)" - o valor de produtos_comp.modelo
aparentemente ja vem com uma "/" no inicio (dado gravado assim). Corrige
removendo qualquer "/" (e espacos) do inicio de cada modelo antes de montar
a string exibida no grid, para virar "CORSA (98/03), CLASSIC(2004/15)".
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta_Pecas.frm"

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


i_dim = find_line_exact("   Dim var_Comp As String     'Compartibilidade")
lines.insert(i_dim + 1, "   Dim sModelo As String")

old_line = '                   var_Comp = var_Comp & r2("modelo") & "(" & r2("ano") & "), "'
i = find_line_exact(old_line, i_dim)

new_lines = [
    '                   sModelo = Trim(r2("modelo"))',
    '                   If Left(sModelo, 1) = "/" Then sModelo = Trim(Mid(sModelo, 2))',
    '                   var_Comp = var_Comp & sModelo & "(" & r2("ano") & "), "',
]

lines[i : i + 1] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - Dim sModelo adicionado, linha {i} substituida por {len(new_lines)} linhas")
