# -*- coding: utf-8 -*-
"""
Adiciona frmAcessorios.Visible e frmSituacao.Visible em cboStatus_Change,
sempre como o oposto de stProdSer.Visible (mesmo padrao de frmParecerCliente):
- A COMECAR (stProdSer=False) -> frmAcessorios/frmSituacao = True
- EM EXECUCAO/AGUARDANDO/TERMINADO (stProdSer=True) -> frmAcessorios/frmSituacao = False
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

# bloco A COMECAR: apos "frmParecerCliente.Visible = True"
i = find_line_exact("   frmParecerCliente.Visible = True", start, end)
lines[i] = lines[i] + (
    "\r\n   frmAcessorios.Visible = True"
    "\r\n   frmSituacao.Visible = True"
)

# bloco EM EXECUCAO/AGUARDANDO: apos "frmParecerCliente.Visible = False" (1a ocorrencia)
end = find_line_exact("End Sub", start)
i = find_line_exact("   frmParecerCliente.Visible = False", i + 1, end)
lines[i] = lines[i] + (
    "\r\n   frmAcessorios.Visible = False"
    "\r\n   frmSituacao.Visible = False"
)

# bloco TERMINADO: apos "frmParecerCliente.Visible = False" (2a ocorrencia)
end = find_line_exact("End Sub", start)
i = find_line_exact("   frmParecerCliente.Visible = False", i + 1, end)
lines[i] = lines[i] + (
    "\r\n   frmAcessorios.Visible = False"
    "\r\n   frmSituacao.Visible = False"
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
