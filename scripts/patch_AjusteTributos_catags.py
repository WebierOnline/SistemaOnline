path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    c = c.replace(old, new, 1)
    print(f'{label} OK')
    return c

# ── 1: Cols = 19 → 21 ──────────────────────────────────────────────────────
content = sub('Cols=21',
    '      .Cols = 19\r\n      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 0\r\n'
    '      .ColWidth(1) = 0\r\n'
    '      .ColWidth(2) = 0\r\n'
    '      .ColWidth(3) = 1300\r\n'
    '      .ColWidth(4) = 4000\r\n'
    '      .ColWidth(5) = 1400\r\n'
    '      .ColWidth(6) = 850\r\n'
    '      .ColWidth(7) = 700\r\n'
    '      .ColWidth(8) = 700\r\n'
    '      .ColWidth(9) = 700\r\n'
    '      .ColWidth(10) = 650\r\n'
    '      .ColWidth(11) = 700\r\n'
    '      .ColWidth(12) = 800\r\n'
    '      .ColWidth(13) = 700\r\n'
    '      .ColWidth(14) = 650\r\n'
    '      .ColWidth(15) = 700\r\n'
    '      .ColWidth(16) = 900\r\n'
    '      .ColWidth(17) = 900\r\n'
    '      .ColWidth(18) = 900\r\n',

    '      .Cols = 21\r\n      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 0\r\n'
    '      .ColWidth(1) = 0\r\n'
    '      .ColWidth(2) = 0\r\n'
    '      .ColWidth(3) = 1300\r\n'
    '      .ColWidth(4) = 4000\r\n'
    '      .ColWidth(5) = 1400\r\n'
    '      .ColWidth(6) = 1400\r\n'
    '      .ColWidth(7) = 1200\r\n'
    '      .ColWidth(8) = 850\r\n'
    '      .ColWidth(9) = 700\r\n'
    '      .ColWidth(10) = 700\r\n'
    '      .ColWidth(11) = 700\r\n'
    '      .ColWidth(12) = 650\r\n'
    '      .ColWidth(13) = 700\r\n'
    '      .ColWidth(14) = 800\r\n'
    '      .ColWidth(15) = 700\r\n'
    '      .ColWidth(16) = 650\r\n'
    '      .ColWidth(17) = 700\r\n'
    '      .ColWidth(18) = 900\r\n'
    '      .ColWidth(19) = 900\r\n'
    '      .ColWidth(20) = 900\r\n',
    content)

# ── 2: Headers cols 6-18 → 8-20 + inserir 6 e 7 ───────────────────────────
content = sub('Headers 6-20',
    '      .TextMatrix(0, 5) = "FABRICANTE"\r\n'
    '      .TextMatrix(0, 6) = "NCM."\r\n'
    '      .TextMatrix(0, 7) = "CFOP."\r\n'
    '      .TextMatrix(0, 8) = "ICMS."\r\n'
    '      .TextMatrix(0, 9) = "ALIQ."\r\n'
    '      .TextMatrix(0, 10) = "PIS"\r\n'
    '      .TextMatrix(0, 11) = "ALIQ."\r\n'
    '      .TextMatrix(0, 12) = "COFINS"\r\n'
    '      .TextMatrix(0, 13) = "ALIQ."\r\n'
    '      .TextMatrix(0, 14) = "IPI"\r\n'
    '      .TextMatrix(0, 15) = "ALIQ."\r\n'
    '      .TextMatrix(0, 16) = "CEST"\r\n'
    '      .TextMatrix(0, 17) = "CL.CBS"\r\n'
    '      .TextMatrix(0, 18) = "CL.IS"\r\n'
    '      \r\n',

    '      .TextMatrix(0, 5) = "FABRICANTE"\r\n'
    '      .TextMatrix(0, 6) = "CATEGORIA"\r\n'
    '      .TextMatrix(0, 7) = "TAGS"\r\n'
    '      .TextMatrix(0, 8) = "NCM."\r\n'
    '      .TextMatrix(0, 9) = "CFOP."\r\n'
    '      .TextMatrix(0, 10) = "ICMS."\r\n'
    '      .TextMatrix(0, 11) = "ALIQ."\r\n'
    '      .TextMatrix(0, 12) = "PIS"\r\n'
    '      .TextMatrix(0, 13) = "ALIQ."\r\n'
    '      .TextMatrix(0, 14) = "COFINS"\r\n'
    '      .TextMatrix(0, 15) = "ALIQ."\r\n'
    '      .TextMatrix(0, 16) = "IPI"\r\n'
    '      .TextMatrix(0, 17) = "ALIQ."\r\n'
    '      .TextMatrix(0, 18) = "CEST"\r\n'
    '      .TextMatrix(0, 19) = "CL.CBS"\r\n'
    '      .TextMatrix(0, 20) = "CL.IS"\r\n'
    '      \r\n',
    content)

# ── 3: DataFill cols 6-18 → 8-20 + inserir 6 (var_cat) e 7 (var_tags) ─────
content = sub('DataFill 6-20',
    '            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))\r\n'
    '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_NCM"))\r\n'
    '            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_CFOP"))\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_ICMS"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = Format$(ValidateNull(rTabela("var_ICMSALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_PIS"))\r\n'
    '            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("var_PISALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("var_COFINS"))\r\n'
    '            .TextMatrix(.rows - 1, 13) = Format$(ValidateNull(rTabela("var_COFINSALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela("var_IPI"))\r\n'
    '            .TextMatrix(.rows - 1, 15) = Format$(ValidateNull(rTabela("var_IPIALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("var_CEST"))\r\n'
    '            .TextMatrix(.rows - 1, 17) = ValidateNull(rTabela("var_CBS"))\r\n'
    '            .TextMatrix(.rows - 1, 18) = ValidateNull(rTabela("var_IS"))\r\n',

    '            .TextMatrix(.rows - 1, 5) = ValidateNull(rTabela("var_fab"))\r\n'
    '            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela("var_cat"))\r\n'
    '            .TextMatrix(.rows - 1, 7) = ValidateNull(rTabela("var_tags"))\r\n'
    '            .TextMatrix(.rows - 1, 8) = ValidateNull(rTabela("var_NCM"))\r\n'
    '            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("var_CFOP"))\r\n'
    '            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("var_ICMS"))\r\n'
    '            .TextMatrix(.rows - 1, 11) = Format$(ValidateNull(rTabela("var_ICMSALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela("var_PIS"))\r\n'
    '            .TextMatrix(.rows - 1, 13) = Format$(ValidateNull(rTabela("var_PISALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela("var_COFINS"))\r\n'
    '            .TextMatrix(.rows - 1, 15) = Format$(ValidateNull(rTabela("var_COFINSALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 16) = ValidateNull(rTabela("var_IPI"))\r\n'
    '            .TextMatrix(.rows - 1, 17) = Format$(ValidateNull(rTabela("var_IPIALIQ")), ocMONEY)\r\n'
    '            .TextMatrix(.rows - 1, 18) = ValidateNull(rTabela("var_CEST"))\r\n'
    '            .TextMatrix(.rows - 1, 19) = ValidateNull(rTabela("var_CBS"))\r\n'
    '            .TextMatrix(.rows - 1, 20) = ValidateNull(rTabela("var_IS"))\r\n',
    content)

# ── 4: SELECT cmdLocalizar — adicionar var_cat e var_tags ──────────────────
content = sub('SELECT cmdLocalizar cat/tags',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS " & _\r\n'
    '      "FROM produtos " & _\r\n'
    '      "WHERE (produtos.ativo = 1) "',

    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags " & _\r\n'
    '      "FROM produtos " & _\r\n'
    '      "WHERE (produtos.ativo = 1) "',
    content)

# ── 5: SELECT MostrarCriterios — adicionar var_cat e var_tags ──────────────
content = sub('SELECT MostrarCriterios cat/tags',
    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS " & _\r\n'
    '      "FROM produtos " & var_Criterio',

    'produtos.IPIALIQ AS var_IPIALIQ, produtos.CEST AS var_CEST, produtos.cClassTrib AS var_CBS, produtos.cClassTrib_IS AS var_IS, produtos.categoria AS var_cat, produtos.TAGS AS var_tags " & _\r\n'
    '      "FROM produtos " & var_Criterio',
    content)

# ── 6: cmdAtualizar — reindexar tudo + inserir categoria e TAGS ─────────────
content = sub('cmdAtualizar reindex',
    '   sSQL = "UPDATE produtos SET " & _\r\n'
    '      "cod_barra = \'" & Grid.TextMatrix(i, 3) & "\', " & _\r\n'
    '      "descricao = \'" & Grid.TextMatrix(i, 4) & "\', " & _\r\n'
    '      "NCM = \'" & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "CFOP = " & Grid.TextMatrix(i, 7) & ", " & _\r\n'
    '      "ICMSCST = \'" & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "ICMSaliq = " & Replace(CDbl(Grid.TextMatrix(i, 9)), ",", ".") & ", " & _\r\n'
    '      "pisCST = \'" & Grid.TextMatrix(i, 10) & "\', " & _\r\n'
    '      "pisaliq = " & Replace(CDbl(Grid.TextMatrix(i, 11)), ",", ".") & ", " & _\r\n'
    '      "cofinsCST = \'" & Grid.TextMatrix(i, 12) & "\', " & _\r\n'
    '      "cofinsaliq = " & Replace(CDbl(Grid.TextMatrix(i, 13)), ",", ".") & ", " & _\r\n'
    '      "ipiCST = \'" & Grid.TextMatrix(i, 14) & "\', " & _\r\n'
    '      "ipialiq = " & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & ", " & _\r\n'
    '      "cest = \'" & Grid.TextMatrix(i, 16) & "\', " & _\r\n'
    '      "cClassTrib = \'" & Grid.TextMatrix(i, 17) & "\', " & _\r\n'
    '      "cClassTrib_IS = \'" & Grid.TextMatrix(i, 18) & "\' " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',

    '   sSQL = "UPDATE produtos SET " & _\r\n'
    '      "cod_barra = \'" & Grid.TextMatrix(i, 3) & "\', " & _\r\n'
    '      "descricao = \'" & Grid.TextMatrix(i, 4) & "\', " & _\r\n'
    '      "categoria = \'" & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "TAGS = \'" & Grid.TextMatrix(i, 7) & "\', " & _\r\n'
    '      "NCM = \'" & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "CFOP = " & Grid.TextMatrix(i, 9) & ", " & _\r\n'
    '      "ICMSCST = \'" & Grid.TextMatrix(i, 10) & "\', " & _\r\n'
    '      "ICMSaliq = " & Replace(CDbl(Grid.TextMatrix(i, 11)), ",", ".") & ", " & _\r\n'
    '      "pisCST = \'" & Grid.TextMatrix(i, 12) & "\', " & _\r\n'
    '      "pisaliq = " & Replace(CDbl(Grid.TextMatrix(i, 13)), ",", ".") & ", " & _\r\n'
    '      "cofinsCST = \'" & Grid.TextMatrix(i, 14) & "\', " & _\r\n'
    '      "cofinsaliq = " & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & ", " & _\r\n'
    '      "ipiCST = \'" & Grid.TextMatrix(i, 16) & "\', " & _\r\n'
    '      "ipialiq = " & Replace(CDbl(Grid.TextMatrix(i, 17)), ",", ".") & ", " & _\r\n'
    '      "cest = \'" & Grid.TextMatrix(i, 18) & "\', " & _\r\n'
    '      "cClassTrib = \'" & Grid.TextMatrix(i, 19) & "\', " & _\r\n'
    '      "cClassTrib_IS = \'" & Grid.TextMatrix(i, 20) & "\' " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',
    content)

# ── 7: Grid_Click loop 6 To 18 → 6 To 20 ──────────────────────────────────
content = sub('GridClick 6-20',
    'For i = 6 To 18\r\n',
    'For i = 6 To 20\r\n',
    content)

# ── 8: txtEdit_LostFocus — reindexar + casos 6 e 7 ─────────────────────────
content = sub('LostFocus reindex',
    'If iCol = 6 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 9 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 11 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 13 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 15 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 16 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 17 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 18 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n',

    'If iCol = 6 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 7 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 8 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 11 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 13 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 15 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 17 Then\r\n'
    '    txtEdit.Text = FormatNumber(txtEdit, 2)\r\n'
    'ElseIf iCol = 18 Then\r\n'
    '    txtEdit.Text = Replace(txtEdit.Text, ".", "")\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 19 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'ElseIf iCol = 20 Then\r\n'
    '    txtEdit.Text = Trim(txtEdit.Text)\r\n'
    'End If\r\n',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('File written successfully.')
