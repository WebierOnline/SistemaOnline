# -*- coding: utf-8 -*-
"""
Adiciona cmdPlaca_Click em OS_Recapadora.frm: abre frmBuscarPlaca (novo
form modal), e se um veiculo for escolhido (USAR ESSE), preenche
CLIENTE/MODELO/ANO/PLACA/KM/COR/CHASSI na aba CADASTRO.
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


anchor = find_line_exact("Private Sub cboCliente_KeyPress(KeyAscii As Integer)")
end_anchor = find_line_exact("End Sub", anchor, anchor + 5)
assert lines[end_anchor + 1] == ""

new_lines = [
    "",
    "Private Sub cmdPlaca_Click()",
    "If vTipoOS <> \"Autom\xf3veis\" And vTipoOS <> \"Motocicletas\" And vTipoOS <> \"Recapadora\" Then",
    "   MsgBox \"Consulta por Placa dispon\xedvel apenas para ve\xedculos!\", vbInformation, \"Aviso do Sistema\"",
    "   Exit Sub",
    "End If",
    "",
    "With frmBuscarPlaca",
    "   .Show vbModal",
    "   If .sPlacaSel <> \"\" Then",
    "      txtCodCliente.Text = .lCodClienteSel",
    "      cboCliente.Text = .sNomeClienteSel & IIf(Trim(.sCelularClienteSel) = \"\", \"\", \"     (\" & Right$(.sCelularClienteSel, 9) & \")\")",
    "      cboModelo.Text = .sModeloSel",
    "      txtAno.Text = .sAnoSel",
    "      txtPlaca.Text = .sPlacaSel",
    "      txtKM.Text = .sKmSel",
    "      cboCor.Text = .sCorSel",
    "      txtChassi.Text = .sChassiSel",
    "   End If",
    "   Unload frmBuscarPlaca",
    "End With",
    "End Sub",
]

lines[end_anchor + 1 : end_anchor + 1] = new_lines

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print(f"OK - cmdPlaca_Click inserido apos linha {end_anchor}")
