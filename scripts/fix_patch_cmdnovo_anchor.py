# -*- coding: utf-8 -*-
"""
Corrige erro do patch anterior: o reset de botoes de servico foi inserido
em cmdImpEntrada1_Click (ancora ambigua) em vez de cmdNovo_Click.
Remove de cmdImpEntrada1_Click e insere corretamente em cmdNovo_Click.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

WRONG_BLOCK = (
    "cmdApagar.Enabled = False\r\n"
    "cmdAdicionarServicosAuto.Enabled = True\r\n"
    "cmdRemoverServicosAuto.Enabled = True\r\n"
    "cmdEditarServicosAuto.Enabled = False\r\n"
    'vCodItemServicoEditando = ""\r\n'
    'cboMecanicoServ.Text = ""\r\n'
    'vCodMecanicoServ = ""'
)
assert text.count(WRONG_BLOCK) == 1, text.count(WRONG_BLOCK)
text = text.replace(WRONG_BLOCK, "cmdApagar.Enabled = False", 1)

# ---------------------------------------------------------------
# Insere corretamente dentro de cmdNovo_Click, apos "cmdApagar.Enabled = False"
# que agora eh a UNICA ocorrencia dentro do range dessa sub.
# ---------------------------------------------------------------
lines = text.split("\r\n")


def find_line_exact(s, start=0, end=None):
    end = end if end is not None else len(lines)
    for i in range(start, end):
        if lines[i] == s:
            return i
    raise SystemExit(f"ERRO: linha exata nao encontrada: {s!r}")


start_novo = find_line_exact("Private Sub cmdNovo_Click()")
end_novo = find_line_exact("End Sub", start_novo)

i = find_line_exact("cmdApagar.Enabled = False", start_novo, end_novo)
lines[i] = (
    lines[i]
    + "\r\n"
    + "cmdAdicionarServicosAuto.Enabled = True\r\n"
    + "cmdRemoverServicosAuto.Enabled = True\r\n"
    + "cmdEditarServicosAuto.Enabled = False\r\n"
    + 'vCodItemServicoEditando = ""\r\n'
    + 'cboMecanicoServ.Text = ""\r\n'
    + 'vCodMecanicoServ = ""'
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
out = out.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

with open(PATH, "wb") as f:
    f.write(out)

print("OK - corrigido")
print("bytes originais:", len(raw), "bytes finais:", len(out))
