# -*- coding: utf-8 -*-
"""
Corrige erro 91 (Object variable or With block variable not set) em
MostrarGrid_OS, na linha "If r.State <> 0 Then r.Close" - reportado pelo
usuario ao abrir o form.

Causa provavel: Form_Load le vTipoOS de sysConfig("TIPO_OS") e chama
MostrarGrid_OS logo em seguida. O If/ElseIf de MostrarGrid_OS so cobre 3
grupos de vTipoOS (Automoveis/Motocicletas/Recapadora,
Informatica/Celular, Comunicacao Visual). Se o valor configurado no banco
(tabela configuracao, config_nome='TIPO_OS') nao bater exatamente com
nenhuma dessas strings (o valor e lido cru, sem Trim/UCase - qualquer
espaco extra ou diferenca de caixa quebra o match), sSQL nunca e montado
e o OpenRecordset falha/retorna algo que deixa r sem State valido.

Fix defensivo: adiciona um Else ao If/ElseIf de vTipoOS que limpa o grid
e sai da sub sem tentar abrir recordset, em vez de estourar erro 91.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

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


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


start, end = find_sub("MostrarGrid_OS")

# a linha "End If" que fecha o If/ElseIf de vTipoOS eh a que vem logo
# antes do comentario "'Debug.Print sSQL"
debug_line = find_line_exact("'Debug.Print sSQL", start, end)
outer_endif = debug_line - 1
assert lines[outer_endif].strip() == "End If", lines[outer_endif]

new_lines = [
    "Else",
    "    FormatarGrid_OS Nothing",
    '    lblQuant.Caption = "QUANTIDADE: " & Format(0, "000")',
    "    Exit Sub",
]
lines[outer_endif:outer_endif] = new_lines

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
