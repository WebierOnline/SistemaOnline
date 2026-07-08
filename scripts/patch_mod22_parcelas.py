# -*- coding: utf-8 -*-
"""
Patch mod22 - Parcelas_Consulta_Produtos.frm (Compartilhado)
Adiciona colunas COD. BARRA (antes DESCRICAO) e COD. PRODUTO (ultima coluna)
no FormatarGrid_Itens, e acrescenta os campos nas SQLs do loadPedidos.
"""
import sys

FILE = r'C:\Projeto\Compartilhado\Forms\Parcelas_Consulta_Produtos.frm'

with open(FILE, 'rb') as f:
    raw = f.read()
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = raw.decode('windows-1252')

changes = []

# ------------------------------------------------------------------
# SQL — PRODUTO query (aparece 2x: count + display) — replace_all
# ------------------------------------------------------------------
old_sql_prod = (
    "pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS \" & _\n"
    "      \"FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto \" & _\n"
)
new_sql_prod = (
    "pedidos_itens.subtotal as var_Subtotal, pedidos_itens.desconto, '' as var_CodOS, ISNULL(produtos.COD_BARRA,'') as var_CodBarra, pedidos_itens.cod_produto as var_CodProd \" & _\n"
    "      \"FROM produtos LEFT JOIN pedidos_itens ON produtos.codigo = pedidos_itens.cod_produto \" & _\n"
)
changes.append((old_sql_prod, new_sql_prod, 'SQL PRODUTO adiciona COD_BARRA e cod_produto', True))

# ------------------------------------------------------------------
# SQL — SERVICO UNION (aparece 2x) — replace_all
# ------------------------------------------------------------------
old_sql_serv = (
    "OS_Servicos_Auto.subtotal as var_Subtotal, OS_Servicos_Auto.desconto, OS_Servicos_Auto.cod_os as var_CodOS \" & _\n"
)
new_sql_serv = (
    "OS_Servicos_Auto.subtotal as var_Subtotal, OS_Servicos_Auto.desconto, OS_Servicos_Auto.cod_os as var_CodOS, '' as var_CodBarra, NULL as var_CodProd \" & _\n"
)
changes.append((old_sql_serv, new_sql_serv, 'SQL SERVICO adiciona campos vazios', True))

# ------------------------------------------------------------------
# SQL — RECEBER a_receber_itens (aparece 2x) — replace_all
# ------------------------------------------------------------------
old_sql_rec1 = (
    "'' as desconto, '' as var_CodOS FROM a_receber_itens WHERE (cod_pedido = "
)
new_sql_rec1 = (
    "'' as desconto, '' as var_CodOS, '' as var_CodBarra, NULL as var_CodProd FROM a_receber_itens WHERE (cod_pedido = "
)
changes.append((old_sql_rec1, new_sql_rec1, 'SQL RECEBER a_receber_itens adiciona campos', True))

# ------------------------------------------------------------------
# SQL — RECEBER pedidos_itens (aparece 2x) — replace_all
# ------------------------------------------------------------------
old_sql_rec2 = (
    "'' as desconto, '' as var_CodOS FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (pedidos_itens.cod_pedido = "
)
new_sql_rec2 = (
    "'' as desconto, '' as var_CodOS, ISNULL(produtos.COD_BARRA,'') as var_CodBarra, pedidos_itens.cod_produto as var_CodProd FROM pedidos_itens INNER JOIN produtos ON pedidos_itens.COD_PRODUTO = produtos.CODIGO WHERE (pedidos_itens.cod_pedido = "
)
changes.append((old_sql_rec2, new_sql_rec2, 'SQL RECEBER pedidos_itens adiciona campos', True))

# ------------------------------------------------------------------
# SQL — ALUGUEL (aparece 2x) — replace_all
# ------------------------------------------------------------------
old_sql_alug = (
    "Aluguel_Cadastro_Itens.SUBTOTAL as var_Subtotal, '' as var_CodOS \" & _\n"
)
new_sql_alug = (
    "Aluguel_Cadastro_Itens.SUBTOTAL as var_Subtotal, '' as var_CodOS, '' as var_CodBarra, NULL as var_CodProd \" & _\n"
)
changes.append((old_sql_alug, new_sql_alug, 'SQL ALUGUEL adiciona campos vazios', True))

# ------------------------------------------------------------------
# FormatarGrid_Itens — Cols + ColWidths
# ------------------------------------------------------------------
old_grid_cols = (
    "      .Cols = 8\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 950\n"
    "      .ColWidth(2) = 5700\n"
    "      .ColWidth(3) = 1000\n"
    "      .ColWidth(4) = 900\n"
    "      .ColWidth(5) = 1100\n"
    "      .ColWidth(6) = 900\n"
    "      .ColWidth(7) = 1000\n"
)
new_grid_cols = (
    "      .Cols = 10\n"
    "      .rows = 2\n"
    "      \n"
    "      .ColWidth(0) = 0\n"
    "      .ColWidth(1) = 950\n"
    "      .ColWidth(2) = 1200\n"
    "      .ColWidth(3) = 4500\n"
    "      .ColWidth(4) = 1000\n"
    "      .ColWidth(5) = 900\n"
    "      .ColWidth(6) = 1100\n"
    "      .ColWidth(7) = 900\n"
    "      .ColWidth(8) = 1000\n"
    "      .ColWidth(9) = 1200\n"
)
changes.append((old_grid_cols, new_grid_cols, 'FormatarGrid_Itens Cols+ColWidths', False))

# ------------------------------------------------------------------
# FormatarGrid_Itens — cabecalhos
# ------------------------------------------------------------------
old_grid_hdr = (
    "      .TextMatrix(0, 1) = \"TIPO\"\n"
    "      .TextMatrix(0, 2) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 3) = \"PREÇO\"\n"
    "      .TextMatrix(0, 4) = \"QUANT\"\n"
    "      .TextMatrix(0, 5) = \"SUBTOTAL\"\n"
    "      .TextMatrix(0, 6) = \"DESC\"\n"
    "      .TextMatrix(0, 7) = \"TOTAL\"\n"
)
new_grid_hdr = (
    "      .TextMatrix(0, 1) = \"TIPO\"\n"
    "      .TextMatrix(0, 2) = \"CÓD. BARRA\"\n"
    "      .TextMatrix(0, 3) = \"DESCRIÇÃO\"\n"
    "      .TextMatrix(0, 4) = \"PREÇO\"\n"
    "      .TextMatrix(0, 5) = \"QUANT\"\n"
    "      .TextMatrix(0, 6) = \"SUBTOTAL\"\n"
    "      .TextMatrix(0, 7) = \"DESC\"\n"
    "      .TextMatrix(0, 8) = \"TOTAL\"\n"
    "      .TextMatrix(0, 9) = \"CÓD. PRODUTO\"\n"
)
changes.append((old_grid_hdr, new_grid_hdr, 'FormatarGrid_Itens cabecalhos', False))

# ------------------------------------------------------------------
# FormatarGrid_Itens — loop de dados
# ------------------------------------------------------------------
old_grid_data = (
    "            .TextMatrix(.rows - 1, 1) = rTabela(\"tipo_item\")\n"
    "            \n"
    "            If tipoEmpresa = 4 Then\n"
    "            .TextMatrix(.rows - 1, 2) = rTabela(\"var_desc\") & \" /  \" & rTabela(\"var_tam\") & \" / \" & rTabela(\"var_fab\")\n"
    "            Else\n"
    "            .TextMatrix(.rows - 1, 2) = rTabela(\"var_desc\") & \" /  \" & ValidateNull(rTabela(\"var_fab\"))\n"
    "            End If\n"
    "            \n"
    "            .TextMatrix(.rows - 1, 3) = Format(rTabela(\"preco\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 4) = rTabela(\"quantidade\")\n"
    "            .TextMatrix(.rows - 1, 5) = Format(rTabela(\"var_Subtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"desconto\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"total\"), ocMONEY)\n"
)
new_grid_data = (
    "            .TextMatrix(.rows - 1, 1) = rTabela(\"tipo_item\")\n"
    "            .TextMatrix(.rows - 1, 2) = ValidateNull(rTabela(\"var_CodBarra\"))\n"
    "            \n"
    "            If tipoEmpresa = 4 Then\n"
    "            .TextMatrix(.rows - 1, 3) = rTabela(\"var_desc\") & \" /  \" & rTabela(\"var_tam\") & \" / \" & rTabela(\"var_fab\")\n"
    "            Else\n"
    "            .TextMatrix(.rows - 1, 3) = rTabela(\"var_desc\") & \" /  \" & ValidateNull(rTabela(\"var_fab\"))\n"
    "            End If\n"
    "            \n"
    "            .TextMatrix(.rows - 1, 4) = Format(rTabela(\"preco\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 5) = rTabela(\"quantidade\")\n"
    "            .TextMatrix(.rows - 1, 6) = Format(rTabela(\"var_Subtotal\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 7) = Format(rTabela(\"desconto\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 8) = Format(rTabela(\"total\"), ocMONEY)\n"
    "            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela(\"var_CodProd\"))\n"
)
changes.append((old_grid_data, new_grid_data, 'FormatarGrid_Itens loop de dados', False))

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
