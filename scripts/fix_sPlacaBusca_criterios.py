# -*- coding: utf-8 -*-
"""
OS_Consulta.frm: no branch de sPlacaBusca (Form_Load), acrescenta
cboConsultaCriterios.Text = "TODOS" + AtualizarCamposCriterios para
esconder os controles de data/nome (calendarios, radio Entrada/
Termino/Previsao) que ficavam visiveis/com "xxx" por nao terem sido
inicializados nesse fluxo.
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


i = find_line_exact('   sPlacaBusca = ""')
assert lines[i + 1] == "   MostrarGrid_OS_Refinado"

new_lines = [
    '   sPlacaBusca = ""',
    '   cboConsultaCriterios.Text = "TODOS"',
    "   AtualizarCamposCriterios",
    "   MostrarGrid_OS_Refinado",
]
lines[i : i + 2] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - cboConsultaCriterios/AtualizarCamposCriterios adicionados ao branch sPlacaBusca")
