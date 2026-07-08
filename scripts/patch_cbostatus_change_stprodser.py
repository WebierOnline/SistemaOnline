# -*- coding: utf-8 -*-
"""
Adiciona reacao em tempo real do stProdSer em cboStatus_Change:
"À COMEÇAR" -> oculto; qualquer outro status -> visivel.
Antes, stProdSer so era atualizado ao abrir/salvar a OS
(cmdEditarOS_Click/cmdAlterar_Click/cmdNovo_Click), nao quando o usuario
so trocava o combo de status sem salvar.
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


start = find_line_exact("Private Sub cboStatus_Change()")
end = find_line_exact("End Sub", start)

# 1a ocorrencia de "frmServicos.Enabled = True" -> branch "A COMECAR" -> oculta
i1 = find_line_exact("   frmServicos.Enabled = True", start, end)
lines[i1] = lines[i1] + "\r\n   stProdSer.Visible = False"

# recalcula o range (uma linha foi inserida) para achar a 2a ocorrencia
end = find_line_exact("End Sub", start)
i2 = find_line_exact("   frmServicos.Enabled = True", i1 + 2, end)
lines[i2] = lines[i2] + "\r\n   stProdSer.Visible = True"

end = find_line_exact("End Sub", start)
i3 = find_line_exact("   frmServicos.Enabled = False", start, end)
lines[i3] = lines[i3] + "\r\n   stProdSer.Visible = True"

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
