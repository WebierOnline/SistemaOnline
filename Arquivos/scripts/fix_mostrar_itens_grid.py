path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: substitui o bloco SQL em Mostrar_Itens ----
old1 = (
"If txtCodEntrada.Text = \"\" Then\r\n"
"   sSQL_Itens = \"SELECT * FROM produtos_entrada_itens WHERE 1 = 0;\"\r\n"
"   Set r = dbData.OpenRecordset(sSQL_Itens)\r\n"
"Else\r\n"
"   sSQL_Itens = \"SELECT produtos_entrada_itens.*, produtos_entrada_itens.codigo as varCod, produtos.COD_BARRA as var_CodBarra, produtos.REF as var_REF, produtos.codigo, produtos.descricao as varDesc, produtos.tamanho, produtos.fabricante \" & _\r\n"
"          \" FROM produtos INNER JOIN produtos_entrada_itens ON produtos.codigo = produtos_entrada_itens.CodigoProduto \" & _\r\n"
"          \" WHERE (codigo_entrada = \" & txtCodEntrada.Text & \") ORDER BY varCod;\"\r\n"
"   Set r = dbData.OpenRecordset(sSQL_Itens)\r\n"
"End If\r\n"
)
new1 = (
"Dim sSel As String\r\n"
"sSel = \"SELECT pei.codigo AS varCod, pei.codigo_entrada, pei.CodigoProduto, \" & _\r\n"
"       \"pei.NomeProduto, ISNULL(pei.EAN,'') AS ean, \" & _\r\n"
"       \"ISNULL(pei.UnidadeTributavel,'') AS UnidadeTributavel, \" & _\r\n"
"       \"ISNULL(pei.NCM,'') AS NCM, ISNULL(pei.CFOP,0) AS CFOP, \" & _\r\n"
"       \"ISNULL(pei.CST,'') AS CST, ISNULL(pei.pICMSST,0) AS pICMSST, \" & _\r\n"
"       \"ISNULL(pei.PISCST,'') AS PISCST, ISNULL(pei.PISpPIS,0) AS PISpPIS, \" & _\r\n"
"       \"ISNULL(pei.COFINSCST,'') AS COFINSCST, \" & _\r\n"
"       \"ISNULL(pei.COFINSpCOFINS,0) AS COFINSpCOFINS, \" & _\r\n"
"       \"ISNULL(pei.QuantidadeTributavel,0) AS QuantidadeTributavel, \" & _\r\n"
"       \"ISNULL((SELECT TOP 1 pp.VALOR_VV FROM Produtos_Precos pp \" & _\r\n"
"       \"        WHERE pp.COD_PRODUTO = pei.CodigoProduto \" & _\r\n"
"       \"          AND pp.COD_ENTRADA = pei.codigo_entrada),0) AS VALOR_VV, \" & _\r\n"
"       \"pei.itemxml, p.COD_BARRA AS var_CodBarra, \" & _\r\n"
"       \"ISNULL(p.REF,'') AS var_REF, ISNULL(p.tamanho,'') AS tamanho \" & _\r\n"
"       \"FROM produtos p INNER JOIN produtos_entrada_itens pei \" & _\r\n"
"       \"ON p.codigo = pei.CodigoProduto\"\r\n"
"If txtCodEntrada.Text = \"\" Then\r\n"
"   sSQL_Itens = sSel & \" WHERE 1=0\"\r\n"
"   Set r = dbData.OpenRecordset(sSQL_Itens)\r\n"
"Else\r\n"
"   sSQL_Itens = sSel & \" WHERE pei.codigo_entrada = \" & txtCodEntrada.Text & \" ORDER BY varCod\"\r\n"
"   Set r = dbData.OpenRecordset(sSQL_Itens)\r\n"
"End If\r\n"
)

# ---- R2: ColWidths das colunas 6..17 ----
old2 = (
"    .ColWidth(6) = 700\r\n"
"    .ColWidth(7) = 850\r\n"
"    .ColWidth(8) = 650\r\n"
"    .ColWidth(9) = 850\r\n"
"    .ColWidth(10) = 650\r\n"
"    .ColWidth(11) = 850\r\n"
"    .ColWidth(12) = 650\r\n"
"    .ColWidth(13) = 850\r\n"
"    .ColWidth(14) = 650\r\n"
"    .ColWidth(15) = 850\r\n"
"    .ColWidth(16) = 0\r\n"
"    .ColWidth(17) = 500\r\n"
)
new2 = (
"    .ColWidth(6) = 700\r\n"
"    .ColWidth(7) = 900\r\n"
"    .ColWidth(8) = 600\r\n"
"    .ColWidth(9) = 600\r\n"
"    .ColWidth(10) = 700\r\n"
"    .ColWidth(11) = 700\r\n"
"    .ColWidth(12) = 700\r\n"
"    .ColWidth(13) = 800\r\n"
"    .ColWidth(14) = 900\r\n"
"    .ColWidth(15) = 1000\r\n"
"    .ColWidth(16) = 900\r\n"
"    .ColWidth(17) = 500\r\n"
)

# ---- R3: cabecalhos das colunas 6..17 ----
old3 = (
"    .TextMatrix(0, 6) = \"QTDE\"\r\n"
"    .TextMatrix(0, 7) = \"CUSTO\"\r\n"
"    .TextMatrix(0, 8) = \"% VV\"\r\n"
"    .TextMatrix(0, 9) = \"VALOR\"\r\n"
"    .TextMatrix(0, 10) = \"% VP \"\r\n"
"    .TextMatrix(0, 11) = \"VALOR\"\r\n"
"    .TextMatrix(0, 12) = \"% AV\"\r\n"
"    .TextMatrix(0, 13) = \"VALOR\"\r\n"
"    .TextMatrix(0, 14) = \"% AP\"\r\n"
"    .TextMatrix(0, 15) = \"VALOR\"\r\n"
"    .TextMatrix(0, 16) = \"SUBTOTAL\"\r\n"
"    .TextMatrix(0, 17) = \"ITEMXML\"\r\n"
)
new3 = (
"    .TextMatrix(0, 6) = \"UNID_TRIB\"\r\n"
"    .TextMatrix(0, 7) = \"NCM\"\r\n"
"    .TextMatrix(0, 8) = \"CFOP\"\r\n"
"    .TextMatrix(0, 9) = \"CST\"\r\n"
"    .TextMatrix(0, 10) = \"% ICMSST\"\r\n"
"    .TextMatrix(0, 11) = \"PISCST\"\r\n"
"    .TextMatrix(0, 12) = \"% PIS\"\r\n"
"    .TextMatrix(0, 13) = \"COFINSCST\"\r\n"
"    .TextMatrix(0, 14) = \"% COFINS\"\r\n"
"    .TextMatrix(0, 15) = \"QTDE TRIB\"\r\n"
"    .TextMatrix(0, 16) = \"VALOR VV\"\r\n"
"    .TextMatrix(0, 17) = \"ITEMXML\"\r\n"
)

# ---- R4: populacao das colunas 6..17 no loop ----
old4 = (
"            .TextMatrix(.rows - 1, 6) = ValidateNull(rTabela(\"QuantidadeTributavel\"))\r\n"
"         .TextMatrix(.rows - 1, 7) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 8) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 9) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 10) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 11) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 12) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 13) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 14) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 15) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 16) = \"\"\r\n"
"         .TextMatrix(.rows - 1, 17) = rTabela(\"itemxml\")\r\n"
)
new4 = (
"         .TextMatrix(.rows - 1, 6)  = ValidateNull(rTabela(\"UnidadeTributavel\"))\r\n"
"         .TextMatrix(.rows - 1, 7)  = ValidateNull(rTabela(\"NCM\"))\r\n"
"         .TextMatrix(.rows - 1, 8)  = ValidateNull(rTabela(\"CFOP\"))\r\n"
"         .TextMatrix(.rows - 1, 9)  = ValidateNull(rTabela(\"CST\"))\r\n"
"         .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela(\"pICMSST\"))\r\n"
"         .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela(\"PISCST\"))\r\n"
"         .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela(\"PISpPIS\"))\r\n"
"         .TextMatrix(.rows - 1, 13) = ValidateNull(rTabela(\"COFINSCST\"))\r\n"
"         .TextMatrix(.rows - 1, 14) = ValidateNull(rTabela(\"COFINSpCOFINS\"))\r\n"
"         .TextMatrix(.rows - 1, 15) = ValidateNull(rTabela(\"QuantidadeTributavel\"))\r\n"
"         .TextMatrix(.rows - 1, 16) = Format(ValidateNull(rTabela(\"VALOR_VV\")), \"##,##0.00\")\r\n"
"         .TextMatrix(.rows - 1, 17) = rTabela(\"itemxml\")\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3 found:', data.count(old3))
print('r4 found:', data.count(old4))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3, new3, 1)
data2 = data2.replace(old4, new4, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
