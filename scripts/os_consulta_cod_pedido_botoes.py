# -*- coding: utf-8 -*-
"""
OS_Consulta.frm:
1) Grid ganha coluna oculta 11 com cod_pedido (ja adicionado ao SELECT
   das 19 consultas antes deste script - substituicao em massa).
2) cmdExibirPedidos_Click / cmdExibirParcelas_Click reescritos para
   usar o layout de colunas do Grid deste form (0=COD_OS, 11=cod_pedido
   oculto), chamando Parcelas_Consulta_Produtos.loadPedidos /
   Vendas_Consulta_Geral_Parcelas.loadInformacoes com o cod_pedido da
   linha selecionada.
"""

PATH = r"C:\projeto\OrdemServico\Forms\OS_Consulta.frm"

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


# ---------------------------------------------------------------
# 1) FormatarGrid_OS: coluna oculta 11 = cod_pedido
# ---------------------------------------------------------------
i = find_line_exact('   .FormatString = "^CÓD.|^ENTRADA|^TERMINO|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL"')
lines[i] = '   .FormatString = "^CÓD.|^ENTRADA|^TERMINO|^TECNICO|^FINANC.|^CLIENTE|^TIPO|^FORMA|^VALOR|^DESC.|^TOTAL|"'

j = find_line_exact("   .ColWidth(10) = 850", i, i + 20)
lines[j] = "   .ColWidth(10) = 850\r\n   .ColWidth(11) = 0"

k = find_line_exact('         .TextMatrix(.Rows - 1, 8) = Format(rTabela("SUBTOTAL"), ocMONEY)', i)
m = find_line_exact('         .TextMatrix(.Rows - 1, 10) = Format(rTabela("TOTAL"), ocMONEY)', k, k + 5)
lines[m] = lines[m] + '\r\n         .TextMatrix(.Rows - 1, 11) = ValidateNull(rTabela("cod_pedido"))'

# ---------------------------------------------------------------
# 2) cmdExibirParcelas_Click / cmdExibirPedidos_Click - reescritos
# ---------------------------------------------------------------
i = find_line_exact("Private Sub cmdExibirParcelas_Click()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub cmdExibirParcelas_Click()",
    "If Grid.Col = 0 Then Exit Sub",
    "   Dim lPedido As Long",
    '   If cboTipo.Text = "POR SERVIÇOS" Then',
    "      If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub",
    "      lPedido = CLng(Grid.TextMatrix(Grid.Row, 11))",
    "      If lPedido = 0 Then Exit Sub",
    "   ElseIf IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then",
    "      lPedido = CLng(Grid.TextMatrix(Grid.Row, 1))",
    "   Else",
    "      Exit Sub",
    "   End If",
    "   Vendas_Consulta_Geral_Parcelas.loadInformacoes lPedido",
    "   Vendas_Consulta_Geral_Parcelas.Show 1",
    "End Sub",
]
assert old == expected, old
novo = [
    "Private Sub cmdExibirParcelas_Click()",
    "If Grid.Row = 0 Then Exit Sub",
    'If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub',
    "If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub",
    "",
    "Dim lPedido As Long",
    "lPedido = CLng(Grid.TextMatrix(Grid.Row, 11))",
    "If lPedido = 0 Then Exit Sub",
    "",
    "Vendas_Consulta_Geral_Parcelas.loadInformacoes lPedido",
    "Vendas_Consulta_Geral_Parcelas.Show 1",
    "End Sub",
]
lines[i : end + 1] = novo

i = find_line_exact("Private Sub cmdExibirPedidos_Click()")
end = find_line_exact("End Sub", i)
old = lines[i : end + 1]
expected = [
    "Private Sub cmdExibirPedidos_Click()",
    'If Grid.TextMatrix(Grid.Row, 1) = "" Then Exit Sub',
    "",
    'If cboTipo.Text = "POR SERVIÇOS" Then',
    "   ' col1 = COD_OS, col11 = COD_PEDIDO (gravado pelo FormatarGrid_Servicos)",
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub",
    "   If CLng(Grid.TextMatrix(Grid.Row, 11)) = 0 Then Exit Sub",
    '   Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 11)), "OS"',
    "Else",
    "   ' POR PRODUTOS: col1 = cod_pedido, col11 = COD_OS (0 = sem OS)",
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub",
    '   If Grid.TextMatrix(Grid.Row, 11) <> "0" Then',
    '      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), "OS"',
    "   Else",
    '      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), "VENDA"',
    "   End If",
    "End If",
    "",
    "Parcelas_Consulta_Produtos.Show 1",
    "End Sub",
]
assert old == expected, old
novo = [
    "Private Sub cmdExibirPedidos_Click()",
    "If Grid.Row = 0 Then Exit Sub",
    'If Grid.TextMatrix(Grid.Row, 0) = "" Then Exit Sub',
    "If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub",
    "If CLng(Grid.TextMatrix(Grid.Row, 11)) = 0 Then Exit Sub",
    "",
    'Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 11)), "OS"',
    "Parcelas_Consulta_Produtos.Show 1",
    "End Sub",
]
lines[i : end + 1] = novo

out_text = "\r\n".join(lines)
out_text = out_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
with open(PATH, "wb") as f:
    f.write(out_text.encode("cp1252"))

print("OK")
