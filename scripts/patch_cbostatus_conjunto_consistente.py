# -*- coding: utf-8 -*-
"""
Estende cboStatus_Change com o conjunto de objetos que ja eram
consistentes (independentes de vTipoOS) em txtCodOS_Change/cmdNovo_Click:
frmParecerCliente, frmGridServicos, frmTotaisGeral, frmTotaisProdServ,
cmdImpEntrada2/Orcamento2/Pedido2.Enabled. frmAcessorios/frmSituacao ficam
de fora (comportamento ambiguo por vTipoOS, usuario pediu para nao mexer).
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

# bloco "A COMECAR" -> oculta
i = find_line_exact("   stProdSer.Visible = False", start, end)
lines[i] = lines[i] + (
    "\r\n   frmParecerCliente.Visible = True"
    "\r\n   frmGridServicos.Visible = False"
    "\r\n   frmTotaisGeral.Visible = False"
    "\r\n   frmTotaisProdServ.Visible = False"
    "\r\n   cmdImpEntrada2.Enabled = False"
    "\r\n   cmdImpOrcamento2.Enabled = False"
    "\r\n   cmdImpPedido2.Enabled = False"
)

# bloco "EM EXECUCAO"/"AGUARDANDO" -> exibe (1a ocorrencia de stProdSer.Visible = True apos o bloco anterior)
end = find_line_exact("End Sub", start)
i = find_line_exact("   stProdSer.Visible = True", i + 1, end)
lines[i] = lines[i] + (
    "\r\n   frmParecerCliente.Visible = False"
    "\r\n   frmGridServicos.Visible = True"
    "\r\n   frmTotaisGeral.Visible = True"
    "\r\n   frmTotaisProdServ.Visible = True"
    "\r\n   cmdImpEntrada2.Enabled = True"
    "\r\n   cmdImpOrcamento2.Enabled = True"
    "\r\n   cmdImpPedido2.Enabled = True"
)

# bloco "TERMINADO" -> exibe (2a ocorrencia de stProdSer.Visible = True)
end = find_line_exact("End Sub", start)
i = find_line_exact("   stProdSer.Visible = True", i + 1, end)
lines[i] = lines[i] + (
    "\r\n   frmParecerCliente.Visible = False"
    "\r\n   frmGridServicos.Visible = True"
    "\r\n   frmTotaisGeral.Visible = True"
    "\r\n   frmTotaisProdServ.Visible = True"
    "\r\n   cmdImpEntrada2.Enabled = True"
    "\r\n   cmdImpOrcamento2.Enabled = True"
    "\r\n   cmdImpPedido2.Enabled = True"
)

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
