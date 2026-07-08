# -*- coding: utf-8 -*-
"""
Debita QUANT_ESTOQUE ao adicionar peca (cmdAdicionarPecas_Click) e devolve
ao remover (cmdRemoverPecas_Click). So nas branches Automoveis/Motocicletas
e Informatica/Celular (vTipoOS real do usuario) - Recapadora/Comunicacao
Visual ficam de fora, mesma restricao aplicada nos patches anteriores.
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


def find_sub(name, start=0):
    s = find_line_exact(f"Private Sub {name}()", start)
    e = find_line_exact("End Sub", s)
    return s, e


# ---------------------------------------------------------------
# 1) cmdAdicionarPecas_Click - debitar estoque
# ---------------------------------------------------------------
start, end = find_sub("cmdAdicionarPecas_Click")
i = find_line_exact("   dbData.Execute sSQL", start, end)
lines[i] = lines[i] + (
    '\r\n   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " '
    '& Replace(CDbl(txtQuantPeca.Text), ",", ".") & " WHERE (codigo = " & txtCodPeca.Text & ");"'
)

# ---------------------------------------------------------------
# 2) cmdRemoverPecas_Click - devolver estoque
# ---------------------------------------------------------------
start2, end2 = find_sub("cmdRemoverPecas_Click")

marker1 = 'dbData.Execute "DELETE FROM pedidos_itens WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 9) & ") AND (cod_pedido = " & txtCodPedido.Text & ");"'
occ = [i for i in range(start2, end2) if lines[i].strip() == marker1]
# 3 ocorrencias: Automoveis/Motocicletas, Informatica/Celular, Comunicacao Visual (nessa ordem).
# So mexer nas 2 primeiras (vTipoOS real do usuario); Comunicacao Visual fica de fora.
assert len(occ) == 3, occ

devolve = (
    '\r\n    dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " '
    '& Replace(CDbl(Grid_Servicos.TextMatrix(Grid_Servicos.Row, 5)), ",", ".") '
    '& " WHERE (codigo = " & Grid_Servicos.TextMatrix(Grid_Servicos.Row, 10) & ");"'
)
for i in occ[:2]:
    lines[i] = lines[i] + devolve

out_text = "\r\n".join(lines)
out = out_text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
