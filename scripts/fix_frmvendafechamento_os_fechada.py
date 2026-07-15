# -*- coding: utf-8 -*-
"""
OS_Recapadora.frm (txtCodOS_Change): no ramo "financeiro ja fechado"
(OS_FINANCEIROABERTO = False), frmVendaFechamento.Visible = False era
setado corretamente e, logo em seguida, sobrescrito para True nos dois
sub-ramos de TIPO_PAGAMENTO (bug de copy-paste do ramo "financeiro
aberto"). Remove as 2 linhas erradas para o painel ficar oculto de
verdade ao carregar uma OS ja fechada, seja por qual caminho for
(frmBuscarPlaca, OS_Consulta, cboLocalizar, etc - todos passam por
txtCodOS_Change).
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
    "    If Not r.BOF Then",
    '        If r("TIPO_PAGAMENTO") = "\xc0 Prazo" Then',
    "            frmVendaFechamento.Visible = True",
    "            cmdFinalizar.Enabled = False",
    "            cmdCancelar.Enabled = False",
    "        Else",
    "            frmVendaFechamento.Visible = True",
    "            cmdFinalizar.Enabled = False",
    "            cmdFinalizar.Enabled = False",
    "        End If",
    "    End If",
]
i = find_line_exact(old_block[0])
for k, l in enumerate(old_block):
    assert lines[i + k] == l, (i + k, repr(lines[i + k]), repr(l))

new_block = [
    "    If Not r.BOF Then",
    '        If r("TIPO_PAGAMENTO") = "\xc0 Prazo" Then',
    "            cmdFinalizar.Enabled = False",
    "            cmdCancelar.Enabled = False",
    "        Else",
    "            cmdFinalizar.Enabled = False",
    "            cmdFinalizar.Enabled = False",
    "        End If",
    "    End If",
]

lines[i : i + len(old_block)] = new_block

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - removidas as 2 linhas 'frmVendaFechamento.Visible = True' erradas (linha {i})")
