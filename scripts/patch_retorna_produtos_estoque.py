# -*- coding: utf-8 -*-
"""
Reescreve Retorna_Produtos_Estoque (chamada por cmdExcluir_Click ao
excluir uma OS inteira) para de fato devolver ao estoque a quantidade de
cada produto vendido naquele pedido. A versao antiga estava inteira
comentada e dependia de um grid (Grid_Pecas) que nao existe mais - alem
disso, cmdExcluir_Click exclui a OS SELECIONADA NA LISTA (Grid_OS), que
pode ser diferente da OS carregada nos campos do formulario, entao usar
Grid_Servicos (que reflete a OS carregada) tambem seria errado.

Nova implementacao: consulta direto pedidos_itens (ja com o cod_pedido
certo, guardado na variavel Public codPedido) via UPDATE...FROM (T-SQL),
antes do DELETE de pedidos_itens/pedidos rodar em cmdExcluir_Click.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Recapadora.frm"

with open(PATH, "rb") as f:
    raw = f.read()

text = raw.decode("cp1252")

OLD = (
    "Private Sub Retorna_Produtos_Estoque()\r\n"
    "'Dim i As Integer\r\n"
    "\r\n"
    "'For i = 1 To Grid_Pecas.Rows - 1\r\n"
    '\'   dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CDbl(Grid_Pecas.TextMatrix(i, 5)), ",", ".") & " WHERE (codigo = " & Grid_Pecas.TextMatrix(i, 2) & ");"\r\n'
    "'Next\r\n"
    "End Sub"
)

NEW = (
    "Private Sub Retorna_Produtos_Estoque()\r\n"
    "dbData.Execute \"UPDATE produtos SET quant_estoque = quant_estoque + pedidos_itens.quantidade \" & _\r\n"
    '   "FROM produtos INNER JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto " & _\r\n'
    '   "WHERE (pedidos_itens.cod_pedido = " & codPedido & ");"\r\n'
    "End Sub"
)

assert text.count(OLD) == 1, "trecho original nao encontrado (ou encontrado mais de uma vez)"
text = text.replace(OLD, NEW, 1)

out = text.encode("cp1252")
with open(PATH, "wb") as f:
    f.write(out)

print("OK - patch aplicado")
