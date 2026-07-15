# -*- coding: utf-8 -*-
"""
OS_Recapadora.frm: implementa Menu_Consulta_Placa_Click (stub vazio ja
criado pelo usuario no IDE) - abre frmBuscarPlaca e, se uma OS for
escolhida (USAR ESSE), carrega ela no OS_Recapadora, espelhando
fielmente o padrao ja usado em Menu_Consulta_OS_Click.
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


old_block = [
    "Private Sub Menu_Consulta_Placa_Click()",
    "",
    "End Sub",
]
i = find_line_exact(old_block[0])
for k, l in enumerate(old_block):
    assert lines[i + k] == l, (i + k, repr(lines[i + k]), repr(l))

new_block = [
    "Private Sub Menu_Consulta_Placa_Click()",
    "If vTipoOS <> \"Autom\xf3veis\" And vTipoOS <> \"Motocicletas\" And vTipoOS <> \"Recapadora\" Then",
    "   MsgBox \"Consulta por Placa dispon\xedvel apenas para ve\xedculos!\", vbInformation, \"Aviso do Sistema\"",
    "   Exit Sub",
    "End If",
    "",
    "frmBuscarPlaca.lCodOSSelecionado = 0",
    "frmBuscarPlaca.Show vbModal",
    "If frmBuscarPlaca.lCodOSSelecionado <> 0 Then",
    "    SSTab1.Tab = 1",
    "    frmSecundario.Enabled = True",
    "    cboStatus.Enabled = True",
    "    cmdGerarEntrada.Enabled = False",
    "    cmdCancelarEntrada.Enabled = False",
    "    cmdAlterar.Enabled = True",
    "    cmdApagar.Enabled = True",
    "    cmdNovo.Enabled = True",
    "    txtCodOS.Text = \"\"",
    "    txtCodOS.Text = frmBuscarPlaca.lCodOSSelecionado",
    "End If",
    "Unload frmBuscarPlaca",
    "End Sub",
]

lines[i : i + len(old_block)] = new_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK - Menu_Consulta_Placa_Click implementado")
