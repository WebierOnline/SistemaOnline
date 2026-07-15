# -*- coding: utf-8 -*-
"""
OS_Consulta.frm: adiciona propriedade publica sPlacaBusca, usada por
frmBuscarPlaca.cmdHistorico_Click para abrir a consulta ja filtrada
(Refinado/Placa) sem passar pelo fluxo padrao (MENSAL) de Form_Load.
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


i_pub = find_line_exact("Public lCodOSSelecionado As Long")
lines.insert(i_pub + 1, "Public sPlacaBusca As String")

old_block = [
    "cboConsultaCriterios.Text = \"MENSAL\"",
    "AtualizarCamposCriterios",
    "",
    "MostrarGrid_OS",
    "End Sub",
]
i_old = find_line_exact(old_block[0])
for k, l in enumerate(old_block):
    assert lines[i_old + k] == l, (i_old + k, repr(lines[i_old + k]), repr(l))

new_block = [
    "If sPlacaBusca <> \"\" Then",
    "   optFiltroRefinado.Value = True",
    "   chkPlaca.Value = 1",
    "   chkChassi.Value = 0",
    "   txtFiltroRefinado.Text = sPlacaBusca",
    "   sPlacaBusca = \"\"",
    "   MostrarGrid_OS_Refinado",
    "Else",
    "   cboConsultaCriterios.Text = \"MENSAL\"",
    "   AtualizarCamposCriterios",
    "   MostrarGrid_OS",
    "End If",
    "End Sub",
]

lines[i_old : i_old + len(old_block)] = new_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - sPlacaBusca adicionado + Form_Load ramificado")
