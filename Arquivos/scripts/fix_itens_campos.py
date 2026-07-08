"""
Renomeia campos de produtos_entrada_itens em todos os forms VB6:
  CODIGO_PRODUTO  -> CodigoProduto
  DESCRICAO       -> NomeProduto
  QUANT           -> QuantidadeTributavel
  (EAN nao renomeia, so tipo no DB)
"""

BASE = 'C:/projeto'

def fix(path, changes, allow_multiple=False):
    data = open(path, 'rb').read().decode('windows-1252')
    changed = False
    for old, new in changes:
        n = data.count(old)
        if n == 0:
            print(f'  SKIP (nao encontrado): {repr(old[:80])}')
        elif n > 1 and not allow_multiple:
            print(f'  SKIP (ambiguo {n}x): {repr(old[:80])}')
        else:
            data = data.replace(old, new)
            changed = True
            print(f'  OK ({n}x): {repr(old[:70])}')
    raw = data.encode('windows-1252')
    raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(path, 'wb').write(raw)
    return changed

# ============================================================
# 1. Entrada_Estoque.Frm
# ============================================================
print('\n--- Entrada_Estoque.Frm ---')
fix(BASE + '/OnlineCommerce/Forms/Entrada_Estoque.Frm', [

    # INSERT col list
    (
        '"codigo_produto, " & _\r\n'
        '   "descricao, " & _\r\n'
        '   "quant, " & _\r\n'
        '   "ean, itemxml ) VALUES (',
        '"CodigoProduto, " & _\r\n'
        '   "NomeProduto, " & _\r\n'
        '   "QuantidadeTributavel, " & _\r\n'
        '   "ean, itemxml ) VALUES ('
    ),

    # Mostrar_Itens: remove colunas deletadas e corrige JOIN
    (
        '"SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod,  produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante, produtos_entrada_itens.CUSTO as var_custo, (produtos_entrada_itens.CUSTO * produtos_entrada_itens.QUANT) as varTotalCustoItem, produtos_entrada_itens.VALOR_VV as var_venda  " & _\r\n'
        '          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CODIGO_PRODUTO " & _\r\n'
        '          " WHERE (codigo_entrada = " & txtCodEntrada.Text & ") ORDER BY varCod;"',
        '"SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod, produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante " & _\r\n'
        '          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CodigoProduto " & _\r\n'
        '          " WHERE (codigo_entrada = " & txtCodEntrada.Text & ") ORDER BY varCod;"'
    ),

    # FormatarGrid_Itens: renomeia leituras e esvazia colunas deletadas
    (
        '            .TextMatrix(.rows - 1, 4) = Format$(rTabela("codigo_produto"), "0000")\r\n'
        '         If tipoEmpresa = 4 Then\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("descricao") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")\r\n'
        '         Else\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("descricao")\r\n'
        '         End If\r\n'
        '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("quant"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = Format$(rTabela("custo"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 9) = Format$(rTabela("VALOR_VV"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 11) = Format$(rTabela("VALOR_VP"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 13) = Format$(rTabela("VALOR_AV"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 15) = Format$(rTabela("VALOR_AP"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 16) = Format$(rTabela("varTotalCustoItem"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 17) = rTabela("itemxml")',
        '            .TextMatrix(.rows - 1, 4) = Format$(rTabela("CodigoProduto"), "0000")\r\n'
        '         If tipoEmpresa = 4 Then\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")\r\n'
        '         Else\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto")\r\n'
        '         End If\r\n'
        '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QuantidadeTributavel"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = ""\r\n'
        '         .TextMatrix(.rows - 1, 8) = ""\r\n'
        '         .TextMatrix(.rows - 1, 9) = ""\r\n'
        '         .TextMatrix(.rows - 1, 10) = ""\r\n'
        '         .TextMatrix(.rows - 1, 11) = ""\r\n'
        '         .TextMatrix(.rows - 1, 12) = ""\r\n'
        '         .TextMatrix(.rows - 1, 13) = ""\r\n'
        '         .TextMatrix(.rows - 1, 14) = ""\r\n'
        '         .TextMatrix(.rows - 1, 15) = ""\r\n'
        '         .TextMatrix(.rows - 1, 16) = ""\r\n'
        '         .TextMatrix(.rows - 1, 17) = rTabela("itemxml")'
    ),

    # MostrarValorVenda
    (
        'WHERE (codigo_produto = " & txtCodProduto & ") ORDER BY codigo DESC;"',
        'WHERE (CodigoProduto = " & txtCodProduto & ") ORDER BY codigo DESC;"'
    ),
])

# ============================================================
# 2. Produtos_Entrada.frm
# ============================================================
print('\n--- Produtos_Entrada.frm ---')
fix(BASE + '/OnlineCommerce/Forms/Produtos_Entrada.frm', [

    # Mostrar_Itens: remove colunas deletadas e corrige JOIN
    (
        '"SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod,  produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante, produtos_entrada_itens.CUSTO as var_custo, (produtos_entrada_itens.CUSTO * produtos_entrada_itens.QUANT) as varTotalCustoItem, produtos_entrada_itens.VALOR_VV as var_venda  " & _\r\n'
        '          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CODIGO_PRODUTO " & _\r\n'
        '          " WHERE (codigo_entrada = " & txtCodigo.Text & ") ORDER BY varDesc, TAMANHO, REF;"',
        '"SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod, produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante " & _\r\n'
        '          " FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CodigoProduto " & _\r\n'
        '          " WHERE (codigo_entrada = " & txtCodigo.Text & ") ORDER BY varDesc, TAMANHO, REF;"'
    ),

    # Mostrar_Total: custo_compra nao existe mais
    (
        '"SELECT ISNULL(SUM(custo_compra * quant), 0) AS var_soma_custo FROM produtos_entrada_itens WHERE (codigo_entrada = " & txtCodigo.Text & ");"',
        '"SELECT 0 AS var_soma_custo;"'
    ),

    # Consulta saldo inicial: quant + codigo_produto
    (
        '"(SELECT ISNULL(SUM(produtos_entrada_itens.quant), 0) FROM produtos_entrada_itens " & _\r\n'
        '         "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _\r\n'
        '         "WHERE (codigo_produto = produtos.codigo) AND (produtos_entrada.data_entrada < CONVERT(DATETIME, \'" & Format(dIni, ocDATA) & "\', 103))) - " & _',
        '"(SELECT ISNULL(SUM(produtos_entrada_itens.QuantidadeTributavel), 0) FROM produtos_entrada_itens " & _\r\n'
        '         "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _\r\n'
        '         "WHERE (CodigoProduto = produtos.codigo) AND (produtos_entrada.data_entrada < CONVERT(DATETIME, \'" & Format(dIni, ocDATA) & "\', 103))) - " & _'
    ),

    # Consulta total diario: quant + codigo_produto
    (
        '"(SELECT ISNULL(SUM(produtos_entrada_itens.quant), 0) FROM produtos_entrada_itens " & _\r\n'
        '            "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _\r\n'
        '            "WHERE (codigo_produto = produtos.codigo) AND (produtos_entrada.data_entrada = CONVERT(DATETIME, \'" & Format$(DIA, ocDATA) & "\', 103))) AS total_entrada, " & _',
        '"(SELECT ISNULL(SUM(produtos_entrada_itens.QuantidadeTributavel), 0) FROM produtos_entrada_itens " & _\r\n'
        '            "INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo " & _\r\n'
        '            "WHERE (CodigoProduto = produtos.codigo) AND (produtos_entrada.data_entrada = CONVERT(DATETIME, \'" & Format$(DIA, ocDATA) & "\', 103))) AS total_entrada, " & _'
    ),

    # INSERT col list (remove custo/margem deletados)
    (
        '"INSERT INTO produtos_entrada_itens (" & _\r\n'
        '   "codigo, " & _\r\n'
        '   "codigo_entrada, " & _\r\n'
        '   "codigo_produto, " & _\r\n'
        '   "descricao, " & _\r\n'
        '   "quant, " & _\r\n'
        '   "custo, " & _\r\n'
        '   "MARGEM_VV, VALOR_VV, MARGEM_VP, VALOR_VP, MARGEM_AV, VALOR_AV, MARGEM_AP, VALOR_AP ) VALUES (" & _\r\n'
        '   var_COD_ITENS & ", " & txtCodigo.Text & ", " & txtCodProduto.Text & ", \'" & cboDescricao.Text & "\', " & _\r\n'
        '   Replace(CDbl(txtQuant.Text), ",", ".") & ", " & Replace(CCur(txtCusto.Text), ",", ".") & ", " & _\r\n'
        '   Replace(CDbl(varMargemVV), ",", ".") & ", " & Replace(CCur(txtValorVV.Text), ",", ".") & ", " & Replace(CDbl(varMargemVP), ",", ".") & ", " & Replace(CCur(txtValorVP.Text), ",", ".") & ", " & Replace(CDbl(varMargemAV), ",", ".") & ", " & Replace(CCur(txtValorAV.Text), ",", ".") & ", " & Replace(CDbl(varMargemAP), ",", ".") & ", " & Replace(CCur(txtValorAP.Text), ",", ".") & ")"',
        '"INSERT INTO produtos_entrada_itens (" & _\r\n'
        '   "codigo, " & _\r\n'
        '   "codigo_entrada, " & _\r\n'
        '   "CodigoProduto, " & _\r\n'
        '   "NomeProduto, " & _\r\n'
        '   "QuantidadeTributavel ) VALUES (" & _\r\n'
        '   var_COD_ITENS & ", " & txtCodigo.Text & ", " & txtCodProduto.Text & ", \'" & cboDescricao.Text & "\', " & _\r\n'
        '   Replace(CDbl(txtQuant.Text), ",", ".") & ")"'
    ),

    # Consulta PRODUTO: JOIN + WHERE descricao
    (
        '"FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.codigo_produto = produtos.codigo " & _\r\n'
        '          "WHERE (produtos_entrada_itens.descricao = \'" & cboConsDescricao.Text & "\') " & _',
        '"FROM produtos_entrada_itens INNER JOIN produtos_entrada ON produtos_entrada_itens.codigo_entrada = produtos_entrada.codigo INNER JOIN produtos ON produtos_entrada_itens.CodigoProduto = produtos.codigo " & _\r\n'
        '          "WHERE (produtos_entrada_itens.NomeProduto = \'" & cboConsDescricao.Text & "\') " & _'
    ),

    # FormatarGrid_Compras: descricao + QUANT + CUSTO
    (
        '         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("descricao"))\r\n'
        '         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QUANT"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = Format$(rTabela("CUSTO"), ocMONEY)',
        '         .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("NomeProduto"))\r\n'
        '         .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QuantidadeTributavel"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = ""'
    ),

    # FormatarGrid_Itens: codigo_produto, descricao, quant + remove deletados
    (
        '            .TextMatrix(.rows - 1, 4) = Format$(rTabela("codigo_produto"), "0000")\r\n'
        '         If tipoEmpresa = 4 Then\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("descricao") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")\r\n'
        '         Else\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("descricao")\r\n'
        '         End If\r\n'
        '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("quant"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = Format$(rTabela("custo"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 8) = FormatNumber(rTabela("MARGEM_VV"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 9) = Format$(rTabela("VALOR_VV"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("MARGEM_VP"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 11) = Format$(rTabela("VALOR_VP"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("MARGEM_AV"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 13) = Format$(rTabela("VALOR_AV"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("MARGEM_AP"), 2) & "%"\r\n'
        '         .TextMatrix(.rows - 1, 15) = Format$(rTabela("VALOR_AP"), ocMONEY)\r\n'
        '         .TextMatrix(.rows - 1, 16) = Format$(rTabela("varTotalCustoItem"), ocMONEY)',
        '            .TextMatrix(.rows - 1, 4) = Format$(rTabela("CodigoProduto"), "0000")\r\n'
        '         If tipoEmpresa = 4 Then\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto") & " /  " & rTabela("tamanho") & " / " & rTabela("var_ref")\r\n'
        '         Else\r\n'
        '            .TextMatrix(.rows - 1, 5) = rTabela("NomeProduto")\r\n'
        '         End If\r\n'
        '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("QuantidadeTributavel"))\r\n'
        '         .TextMatrix(.rows - 1, 7) = ""\r\n'
        '         .TextMatrix(.rows - 1, 8) = ""\r\n'
        '         .TextMatrix(.rows - 1, 9) = ""\r\n'
        '         .TextMatrix(.rows - 1, 10) = ""\r\n'
        '         .TextMatrix(.rows - 1, 11) = ""\r\n'
        '         .TextMatrix(.rows - 1, 12) = ""\r\n'
        '         .TextMatrix(.rows - 1, 13) = ""\r\n'
        '         .TextMatrix(.rows - 1, 14) = ""\r\n'
        '         .TextMatrix(.rows - 1, 15) = ""\r\n'
        '         .TextMatrix(.rows - 1, 16) = ""'
    ),
], allow_multiple=True)

# ============================================================
# 3. EntradasvsSaidas.frm  (replacements globais, multiplos)
# ============================================================
print('\n--- EntradasvsSaidas.frm ---')
fix(BASE + '/OnlineCommerce/Forms/EntradasvsSaidas.frm', [
    ('produtos_entrada_itens.CODIGO_PRODUTO',  'produtos_entrada_itens.CodigoProduto'),
    ('produtos_entrada_itens.codigo_produto',  'produtos_entrada_itens.CodigoProduto'),
    ('produtos_entrada_itens.QUANT',           'produtos_entrada_itens.QuantidadeTributavel'),
    ('produtos_entrada_itens.descricao as varDesc', 'produtos_entrada_itens.NomeProduto as varDesc'),
    ('produtos_entrada_itens.quant as varQuant',    'produtos_entrada_itens.QuantidadeTributavel as varQuant'),
    ('ORDER BY produtos_entrada_itens.descricao,',  'ORDER BY produtos_entrada_itens.NomeProduto,'),
], allow_multiple=True)

# ============================================================
# 4. Produtos_Cadastro.frm
# ============================================================
print('\n--- Produtos_Cadastro.frm ---')
fix(BASE + '/Compartilhado/Forms/Produtos_Cadastro.frm', [
    (
        'UPDATE produtos_entrada_itens SET descricao = \'" & txtDescricao.Text & "\' WHERE (codigo_produto = " & txtCodigo.Text & ");',
        'UPDATE produtos_entrada_itens SET NomeProduto = \'" & txtDescricao.Text & "\' WHERE (CodigoProduto = " & txtCodigo.Text & ");'
    ),
], allow_multiple=True)

# ============================================================
# 5. Produtos_Saida_Estoque.frm
# ============================================================
print('\n--- Produtos_Saida_Estoque.frm ---')
fix(BASE + '/Compartilhado/Forms/Produtos_Saida_Estoque.frm', [
    (
        '"SELECT ISNULL(SUM(quant), 0) AS total_entradas FROM produtos_entrada_itens WHERE (codigo_produto = " & txtCodProduto & ");"',
        '"SELECT ISNULL(SUM(QuantidadeTributavel), 0) AS total_entradas FROM produtos_entrada_itens WHERE (CodigoProduto = " & txtCodProduto & ");"'
    ),
])

# ============================================================
# 6. PDV.frm
# ============================================================
print('\n--- PDV.frm ---')
fix(BASE + '/PDV/Forms/PDV.frm', [
    # Query 1: valor_vv nao existe mais -> NULL; corrige JOIN
    (
        '"produtos_entrada_itens.valor_vv AS var_venda, produtos_entrada_itens.codigo_produto, produtos.cod_barra, produtos.ativo, " & _\r\n'
        '   "produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.codigo_produto " & _',
        '"NULL AS var_venda, produtos_entrada_itens.CodigoProduto, produtos.cod_barra, produtos.ativo, " & _\r\n'
        '   "produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CodigoProduto " & _'
    ),
    # Query 2: venda nao existe mais -> NULL; corrige JOIN
    (
        '"produtos.descricao AS var_desc, produtos_entrada_itens.venda AS var_venda FROM produtos LEFT JOIN ultimas_entradas ON produtos.codigo = ultimas_entradas.codigo_produto " & _\r\n'
        '      "LEFT JOIN produtos_entrada_itens ON ultimas_entradas.codigo_produto = produtos_entrada_itens.codigo_produto AND ultimas_entradas.ultentrada = produtos_entrada_itens.codigo_entrada " & _',
        '"produtos.descricao AS var_desc, NULL AS var_venda FROM produtos LEFT JOIN ultimas_entradas ON produtos.codigo = ultimas_entradas.codigo_produto " & _\r\n'
        '      "LEFT JOIN produtos_entrada_itens ON ultimas_entradas.codigo_produto = produtos_entrada_itens.CodigoProduto AND ultimas_entradas.ultentrada = produtos_entrada_itens.codigo_entrada " & _'
    ),
    # Query 3: corrige JOIN
    (
        '"produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.codigo_produto " & _\r\n'
        '   "WHERE (produtos.cod_barra = \'" & txtCodBarra.Text & "\') AND (produtos.ativo = 1) ORDER BY produtos_entrada_itens.codigo DESC;"',
        '"produtos_entrada_itens.codigo FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CodigoProduto " & _\r\n'
        '   "WHERE (produtos.cod_barra = \'" & txtCodBarra.Text & "\') AND (produtos.ativo = 1) ORDER BY produtos_entrada_itens.codigo DESC;"'
    ),
])

# ============================================================
# 7. PDV_Consulta.frm  (usa DAO/Access, maiusculas)
# ============================================================
print('\n--- PDV_Consulta.frm ---')
fix(BASE + '/PDV/Forms/PDV_Consulta.frm', [
    ('PRODUTOS_ENTRADA_ITENS.CODIGO_PRODUTO', 'PRODUTOS_ENTRADA_ITENS.CodigoProduto'),
], allow_multiple=True)

# ============================================================
# 8. OS_Automoveis.frm
# ============================================================
print('\n--- OS_Automoveis.frm ---')
fix(BASE + '/OrdemServico/Forms/OS_Automoveis.frm', [
    ('produtos_entrada_itens.codigo_produto = produtos.codigo', 'produtos_entrada_itens.CodigoProduto = produtos.codigo'),
])

# ============================================================
# 9. Ordem_Servicos_Motores.frm
# ============================================================
print('\n--- Ordem_Servicos_Motores.frm ---')
fix(BASE + '/OrdemServico/Forms/Ordem_Servicos_Motores.frm', [
    ('produtos_entrada_itens.codigo_produto = produtos.codigo', 'produtos_entrada_itens.CodigoProduto = produtos.codigo'),
])

# ============================================================
# 10. Ordem_Servicos_Informatica.frm
# ============================================================
print('\n--- Ordem_Servicos_Informatica.frm ---')
fix(BASE + '/OrdemServico/Forms/Ordem_Servicos_Informatica.frm', [
    ('produtos_entrada_itens.codigo_produto = produtos.codigo', 'produtos_entrada_itens.CodigoProduto = produtos.codigo'),
])

print('\nConcluido.')
