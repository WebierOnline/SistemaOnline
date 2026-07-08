# -*- coding: utf-8 -*-
"""
Remove o bloco de MostrarGrid_Servicos que carregava o cod_mecanico do
ultimo servico salvo e pre-preenchia cboMecanicoServ ao reabrir a OS.
Usuario pediu para cboMecanicoServ comecar sempre vazio ao abrir a OS,
preenchendo so quando o usuario escolher um servico (cboServicosAuto).
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


start = find_line_exact("Private Sub MostrarGrid_Servicos()")
end = find_line_exact("End Sub", start)

block_start = find_line_exact("Dim vCodMecServAtual As String", start, end)
# o End If final do bloco eh a ultima linha antes do End Sub
assert lines[end - 1].strip() == "End If", lines[end - 1]

# remove do "Dim vCodMecServAtual" (inclusive) ate o "End If" final (inclusive),
# junto com a linha em branco que precede o bloco
blank_before = block_start - 1
assert lines[blank_before].strip() == "", repr(lines[blank_before])

del lines[blank_before:end]

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - bloco removido")
