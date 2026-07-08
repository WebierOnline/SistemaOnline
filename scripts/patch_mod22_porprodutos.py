# -*- coding: utf-8 -*-
"""
Patch mod22 - Vendas_Consulta_PorProdutos.frm
Adiciona colunas COD. BARRA (antes DESCRICAO) e COD. PRODUTO (ultima coluna)
nos grids FormatarGrid_ProdDetalhado e FormatarGrid_Servicos.
"""
import sys

FILE = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# 1. SQL POR PRODUTOS — acrescenta varCodBarra no SELECT
# ------------------------------------------------------------------
old1 = (
    "ISNULL(OS.COD_OS, 0) AS var_CodOS \" & _\n"
    "        \"FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo LEFT OUTER JOIN OS ON pedidos.COD_PEDIDO = OS.COD_PEDIDO \" & _\n"
)
new1 = (
    "ISNULL(OS.COD_OS, 0) AS var_CodOS, produtos.COD_BARRA as varCodBarra \" & _\n"
    "        \"FROM pedidos_itens INNER JOIN pedidos ON pedidos_itens.cod_pedido = pedidos.cod_pedido INNER JOIN produtos ON pedidos_itens.cod_produto = produtos.codigo LEFT OUTER JOIN OS ON pedidos.COD_PEDIDO = OS.COD_PEDIDO \" & _\n"
)
changes.append((old1, new1, '1 - SQL POR PRODUTOS adiciona varCodBarra', False))

# ------------------------------------------------------------------
# 2. FormatarGrid_ProdDetalhado — Cols + ColWidths
#    Ancora unica: ColWidth(4) = 6000
# ------------------------------------------------------------------
old2 = (
    "      .Cols = 11\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 750\n"
    "      .ColWidth(2) = 900\n"
    "      .ColWidth(3) = 0\n"
    "      .ColWidth(4) = 6000\n"
    "      .ColWidth(5) = 800\n"
    "      .ColWidth(6) = 800\n"
    "      .ColWidth(7) = 800\n"
    "      .ColWidth(8) = 700\n"
    "      .ColWidth(9) = 800\n"
    "      .ColWidth(10) = 0\n"
)
new2 = (
    "      .Cols = 13\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 750\n"
    "      .ColWidth(2) = 900\n"
    "      .ColWidth(3) = 0\n"
    "      .ColWidth(4) = 1200\n"
    "      .ColWidth(5) = 4500\n"
    "      .ColWidth(6) = 800\n"
    "      .ColWidth(7) = 800\n"
    "      .ColWidth(8) = 800\n"
    "      .ColWidth(9) = 700\n"
    "      .ColWidth(10) = 800\n"
    "      .ColWidth(11) = 0\n"
    "      .ColWidth(12) = 1200\n"
)
changes.append((old2, new2, '2 - ProdDetalhado Cols+ColWidths', False))

# ------------------------------------------------------------------
# 3. FormatarGrid_ProdDetalhado — cabecalhos TextMatrix
# ------------------------------------------------------------------
old3 = (
    "      .TextMatrix(0, 1) = \"PEDIDO\"\n"
    "      .TextMatrix(0, 2) = \"DATA\"\n"
    "      .TextMatrix(0, 3) = \"CÓD.PROD.\"\n"
    "      .TextMatrix(0, 4) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 5) = \"VALOR\"\n"
    "      .TextMatrix(0, 6) = \"QTDE\"\n"
    "      .TextMatrix(0, 7) = \"=\"\n"
    "      .TextMatrix(0, 8) = \"DESC.\"\n"
    "      .TextMatrix(0, 9) = \"TOTAL\"\n"
    "      .TextMatrix(0, 10) = \"COD_OS\"\n"
)
new3 = (
    "      .TextMatrix(0, 1) = \"PEDIDO\"\n"
    "      .TextMatrix(0, 2) = \"DATA\"\n"
    "      .TextMatrix(0, 3) = \"CÓD.PROD.\"\n"
    "      .TextMatrix(0, 4) = \"CÓD. BARRA\"\n"
    "      .TextMatrix(0, 5) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 6) = \"VALOR\"\n"
    "      .TextMatrix(0, 7) = \"QTDE\"\n"
    "      .TextMatrix(0, 8) = \"=\"\n"
    "      .TextMatrix(0, 9) = \"DESC.\"\n"
    "      .TextMatrix(0, 10) = \"TOTAL\"\n"
    "      .TextMatrix(0, 11) = \"COD_OS\"\n"
    "      .TextMatrix(0, 12) = \"CÓD. PRODUTO\"\n"
)
changes.append((old3, new3, '3 - ProdDetalhado cabecalhos', False))

# ------------------------------------------------------------------
# 4. FormatarGrid_ProdDetalhado — loop de dados
# ------------------------------------------------------------------
old4 = (
    "            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"varcodped\"), \"000000\")\n"
    "            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"varData\"), \"dd/mm/yy\")\n"
    "            .TextMatrix(.rows - 1, 3) = rTabela(\"varcodprod\")\n"
    "            '.TextMatrix(.Rows - 1, 4) = rTabela(\"vardesc\")\n"
    "            .TextMatrix(.rows - 1, 5) = Format(rTabela(\"varvalor\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 6) = rTabela(\"varquant\")\n"
    "            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"varsubtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"vardesc\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"vartotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 10) = rTabela(\"var_codos\")\n"
    "            \n"
    "            If tipoEmpresa = 4 Then\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"varNome\") & \" /  \" & rTabela(\"vartam\") & \" / \" & rTabela(\"varfab\") & \" /  \" & rTabela(\"varref\")\n"
    "            Else\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"varNome\") & \" /  \" & ValidateNull(rTabela(\"varfab\")) & \" /  \" & rTabela(\"varRef\")\n"
    "            End If\n"
)
new4 = (
    "            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"varcodped\"), \"000000\")\n"
    "            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"varData\"), \"dd/mm/yy\")\n"
    "            .TextMatrix(.rows - 1, 3) = rTabela(\"varcodprod\")\n"
    "            .TextMatrix(.rows - 1, 4) = ValidateNull(rTabela(\"varCodBarra\"))\n"
    "            '.TextMatrix(.Rows - 1, 5) = rTabela(\"vardesc\")\n"
    "            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"varvalor\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 7) = rTabela(\"varquant\")\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"varsubtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"vardesc\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 10) = Format(rTabela(\"vartotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 11) = rTabela(\"var_codos\")\n"
    "            .TextMatrix(.rows - 1, 12) = rTabela(\"varcodprod\")\n"
    "            \n"
    "            If tipoEmpresa = 4 Then\n"
    "            .TextMatrix(.rows - 1, 5) = rTabela(\"varNome\") & \" /  \" & rTabela(\"vartam\") & \" / \" & rTabela(\"varfab\") & \" /  \" & rTabela(\"varref\")\n"
    "            Else\n"
    "            .TextMatrix(.rows - 1, 5) = rTabela(\"varNome\") & \" /  \" & ValidateNull(rTabela(\"varfab\")) & \" /  \" & rTabela(\"varRef\")\n"
    "            End If\n"
)
changes.append((old4, new4, '4 - ProdDetalhado loop de dados', False))

# ------------------------------------------------------------------
# 5. FormatarGrid_ProdDetalhado — SomaGrid
#    Ancora unica: presenca do comentario 'lblEntrada
# ------------------------------------------------------------------
old5 = (
    "   lblQtda.Caption = SomaGrid(Grid, 6)\n"
    "   lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)\n"
    "   'lblEntrada.Caption = Format(0, ocMONEY)\n"
)
new5 = (
    "   lblQtda.Caption = SomaGrid(Grid, 7)\n"
    "   lblTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)\n"
    "   'lblEntrada.Caption = Format(0, ocMONEY)\n"
)
changes.append((old5, new5, '5 - ProdDetalhado SomaGrid cols 6->7 e 9->10', False))

# ------------------------------------------------------------------
# 6. FormatarGrid_Servicos — Cols + ColWidths
#    Ancora unica: ColWidth(4) = 5400
# ------------------------------------------------------------------
old6 = (
    "      .Cols = 11\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 750\n"
    "      .ColWidth(2) = 900\n"
    "      .ColWidth(3) = 0\n"
    "      .ColWidth(4) = 5400\n"
    "      .ColWidth(5) = 800\n"
    "      .ColWidth(6) = 800\n"
    "      .ColWidth(7) = 800\n"
    "      .ColWidth(8) = 700\n"
    "      .ColWidth(9) = 800\n"
    "      .ColWidth(10) = 0\n"
)
new6 = (
    "      .Cols = 13\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 750\n"
    "      .ColWidth(2) = 900\n"
    "      .ColWidth(3) = 0\n"
    "      .ColWidth(4) = 1200\n"
    "      .ColWidth(5) = 4200\n"
    "      .ColWidth(6) = 800\n"
    "      .ColWidth(7) = 800\n"
    "      .ColWidth(8) = 800\n"
    "      .ColWidth(9) = 700\n"
    "      .ColWidth(10) = 800\n"
    "      .ColWidth(11) = 0\n"
    "      .ColWidth(12) = 1200\n"
)
changes.append((old6, new6, '6 - Servicos Cols+ColWidths', False))

# ------------------------------------------------------------------
# 7. FormatarGrid_Servicos — cabecalhos TextMatrix
# ------------------------------------------------------------------
old7 = (
    "      .TextMatrix(0, 1) = \"OS\"\n"
    "      .TextMatrix(0, 2) = \"DATA\"\n"
    "      .TextMatrix(0, 3) = \"\"\n"
    "      .TextMatrix(0, 4) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 5) = \"VALOR\"\n"
    "      .TextMatrix(0, 6) = \"QTDE\"\n"
    "      .TextMatrix(0, 7) = \"=\"\n"
    "      .TextMatrix(0, 8) = \"DESC.\"\n"
    "      .TextMatrix(0, 9) = \"TOTAL\"\n"
    "      .TextMatrix(0, 10) = \"\"\n"
)
new7 = (
    "      .TextMatrix(0, 1) = \"OS\"\n"
    "      .TextMatrix(0, 2) = \"DATA\"\n"
    "      .TextMatrix(0, 3) = \"\"\n"
    "      .TextMatrix(0, 4) = \"CÓD. BARRA\"\n"
    "      .TextMatrix(0, 5) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 6) = \"VALOR\"\n"
    "      .TextMatrix(0, 7) = \"QTDE\"\n"
    "      .TextMatrix(0, 8) = \"=\"\n"
    "      .TextMatrix(0, 9) = \"DESC.\"\n"
    "      .TextMatrix(0, 10) = \"TOTAL\"\n"
    "      .TextMatrix(0, 11) = \"\"\n"
    "      .TextMatrix(0, 12) = \"CÓD. PRODUTO\"\n"
)
changes.append((old7, new7, '7 - Servicos cabecalhos', False))

# ------------------------------------------------------------------
# 8. FormatarGrid_Servicos — loop de dados
# ------------------------------------------------------------------
old8 = (
    "            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"varCodPed\"), \"000000\")\n"
    "            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"varData\"), \"dd/mm/yy\")\n"
    "            .TextMatrix(.rows - 1, 3) = \"\"\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"varNome\")\n"
    "            .TextMatrix(.rows - 1, 5) = Format(rTabela(\"varValor\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 6) = rTabela(\"varQuant\")\n"
    "            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"varSubtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"varDesc\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"varTotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 10) = rTabela(\"var_CodOS\")\n"
)
new8 = (
    "            .TextMatrix(.rows - 1, 1) = Format(rTabela(\"varCodPed\"), \"000000\")\n"
    "            .TextMatrix(.rows - 1, 2) = Format(rTabela(\"varData\"), \"dd/mm/yy\")\n"
    "            .TextMatrix(.rows - 1, 3) = \"\"\n"
    "            .TextMatrix(.rows - 1, 4) = \"\"\n"
    "            .TextMatrix(.rows - 1, 5) = rTabela(\"varNome\")\n"
    "            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"varValor\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 7) = rTabela(\"varQuant\")\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"varSubtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = Format(rTabela(\"varDesc\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 10) = Format(rTabela(\"varTotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 11) = rTabela(\"var_CodOS\")\n"
    "            .TextMatrix(.rows - 1, 12) = \"\"\n"
)
changes.append((old8, new8, '8 - Servicos loop de dados', False))

# ------------------------------------------------------------------
# 9. FormatarGrid_Servicos — SomaGrid
#    Ancora unica: sem comentario 'lblEntrada (termina direto com picAguarde)
# ------------------------------------------------------------------
old9 = (
    "   lblQtda.Caption = SomaGrid(Grid, 6)\n"
    "   lblTotal.Caption = Format(SomaGrid(Grid, 9), ocMONEY)\n"
    "picAguarde.Visible = False\n"
    "End Sub\n"
)
new9 = (
    "   lblQtda.Caption = SomaGrid(Grid, 7)\n"
    "   lblTotal.Caption = Format(SomaGrid(Grid, 10), ocMONEY)\n"
    "picAguarde.Visible = False\n"
    "End Sub\n"
)
changes.append((old9, new9, '9 - Servicos SomaGrid cols 6->7 e 9->10', False))

# ------------------------------------------------------------------
# 10. cmdExibirPedidos — atualiza col 10 -> col 11
# ------------------------------------------------------------------
old10 = (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   ' col1 = COD_OS, col10 = COD_PEDIDO (gravado pelo FormatarGrid_Servicos)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 10)) Then Exit Sub\n"
    "   If CLng(Grid.TextMatrix(Grid.Row, 10)) = 0 Then Exit Sub\n"
    "   Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 10)), \"OS\"\n"
    "Else\n"
    "   ' POR PRODUTOS: col1 = cod_pedido, col10 = COD_OS (0 = sem OS)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub\n"
    "   If Grid.TextMatrix(Grid.Row, 10) <> \"0\" Then\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"OS\"\n"
    "   Else\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"VENDA\"\n"
    "   End If\n"
    "End If\n"
)
new10 = (
    "If cboTipo.Text = \"POR SERVIÇOS\" Then\n"
    "   ' col1 = COD_OS, col11 = COD_PEDIDO (gravado pelo FormatarGrid_Servicos)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 11)) Then Exit Sub\n"
    "   If CLng(Grid.TextMatrix(Grid.Row, 11)) = 0 Then Exit Sub\n"
    "   Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 11)), \"OS\"\n"
    "Else\n"
    "   ' POR PRODUTOS: col1 = cod_pedido, col11 = COD_OS (0 = sem OS)\n"
    "   If Not IsNumeric(Grid.TextMatrix(Grid.Row, 1)) Then Exit Sub\n"
    "   If Grid.TextMatrix(Grid.Row, 11) <> \"0\" Then\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"OS\"\n"
    "   Else\n"
    "      Parcelas_Consulta_Produtos.loadPedidos CLng(Grid.TextMatrix(Grid.Row, 1)), \"VENDA\"\n"
    "   End If\n"
    "End If\n"
)
changes.append((old10, new10, '10 - cmdExibirPedidos col10->col11', False))

# ------------------------------------------------------------------
# Aplicar
# ------------------------------------------------------------------
for old, new, label, replace_all in changes:
    count = text.count(old)
    if replace_all:
        if count == 0:
            print(f'ERRO [{label}]: 0 ocorrencias')
            sys.exit(1)
        text = text.replace(old, new)
        print(f'OK ({count}x): {label}')
    else:
        if count != 1:
            print(f'ERRO [{label}]: encontrado {count} ocorrencias (esperado 1)')
            sys.exit(1)
        text = text.replace(old, new)
        print(f'OK: {label}')

text = text.replace('\r\n', '\n').replace('\r', '\n')
out = text.encode('windows-1252')
out = out.replace(b'\n', b'\r\n')
with open(FILE, 'wb') as f:
    f.write(out)
print('\nArquivo gravado com sucesso.')
