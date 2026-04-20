path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: adiciona leitura de vUnTrib e calcula qtfracao antes das buscas ----
old1 = (
"          'qCom e qTrib: usados para calcular fracionamento quando cEAN(caixa) != cEANTrib(unidade)\r\n"
"          Dim xQCom  As Double\r\n"
"          Dim xQTrib As Double\r\n"
"          xQCom = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qCom\").Item(0).Text, \".\", \",\", UM_A_UM))\r\n"
"          If XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qTrib\").Length > 0 Then\r\n"
"             xQTrib = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qTrib\").Item(0).Text, \".\", \",\", UM_A_UM))\r\n"
"          Else\r\n"
"             xQTrib = xQCom\r\n"
"          End If\r\n"
"\r\n"
"          qtfracao = 1\r\n"
"          CodProduto = 0\r\n"
)

new1 = (
"          'qCom e qTrib: usados para calcular fracionamento quando cEAN(caixa) != cEANTrib(unidade)\r\n"
"          Dim xQCom  As Double\r\n"
"          Dim xQTrib As Double\r\n"
"          Dim xVUnTrib As Double\r\n"
"          xQCom = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qCom\").Item(0).Text, \".\", \",\", UM_A_UM))\r\n"
"          If XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qTrib\").Length > 0 Then\r\n"
"             xQTrib = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qTrib\").Item(0).Text, \".\", \",\", UM_A_UM))\r\n"
"          Else\r\n"
"             xQTrib = xQCom\r\n"
"          End If\r\n"
"          If XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/vUnTrib\").Length > 0 Then\r\n"
"             xVUnTrib = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/vUnTrib\").Item(0).Text, \".\", \",\", UM_A_UM))\r\n"
"          Else\r\n"
"             xVUnTrib = xVUnCom\r\n"
"          End If\r\n"
"\r\n"
"          'fracionamento: qTrib/qCom quando cEANTrib diferente de cEAN (fornecedor usa embalagem)\r\n"
"          If Not Vazio(cEANTrib) And cEANTrib <> cEAN And xQCom > 0 And xQTrib > xQCom Then\r\n"
"             qtfracao = xQTrib / xQCom\r\n"
"          Else\r\n"
"             qtfracao = 1\r\n"
"          End If\r\n"
"          CodProduto = 0\r\n"
)

# ---- R2: substitui Busca3+Busca4 pela nova regra: so vincula por cEANTrib se diferente de cEAN ----
old2 = (
"          '--- Busca 3: Produtos.EAN por cEAN (EAN da embalagem comercial) ---\r\n"
"          If CodProduto = 0 And Not Vazio(cEAN) Then\r\n"
"             CodProduto = SQLExecutaRetorno(\"SELECT Codigo FROM Produtos WHERE EAN = '\" & cEAN & \"'\", \"Codigo\", 0)\r\n"
"          End If\r\n"
"\r\n"
"          '--- Busca 4: Produtos.EAN por cEANTrib ---\r\n"
"          If CodProduto = 0 And Not Vazio(cEANTrib) Then\r\n"
"             CodProduto = SQLExecutaRetorno(\"SELECT Codigo FROM Produtos WHERE EAN = '\" & cEANTrib & \"'\", \"Codigo\", 0)\r\n"
"             If CodProduto > 0 And cEANTrib <> cEAN And xQCom > 0 And xQTrib > xQCom Then\r\n"
"                'cEANTrib diferente de cEAN: fornecedor esta enviando embalagem (caixa)\r\n"
"                'fracionamento = quantas unidades tributaveis ha por unidade comercial\r\n"
"                qtfracao = xQTrib / xQCom\r\n"
"             End If\r\n"
"          End If\r\n"
)

new2 = (
"          '--- Busca 3: Produtos.EAN por cEANTrib (vinculo pela unidade interna) ---\r\n"
"          '   So vincula automaticamente se cEANTrib for diferente de cEAN.\r\n"
"          '   Se cEANTrib = cEAN o fornecedor nao informou unidade separada: aguarda vinculo manual.\r\n"
"          If CodProduto = 0 And Not Vazio(cEANTrib) And cEANTrib <> cEAN Then\r\n"
"             CodProduto = SQLExecutaRetorno(\"SELECT Codigo FROM Produtos WHERE EAN = '\" & cEANTrib & \"'\", \"Codigo\", 0)\r\n"
"          End If\r\n"
)

# ---- R3: CustoUnitario no UPDATE usa xVUnTrib em vez de xVUnCom/qtfracao ----
old3 = (
"                       \"CustoUnitario = \" & FSQL(xVUnCom / qtfracao) & \", \" & _\r\n"
"                       \"DataAtualizacao = GETDATE(), \"\r\n"
)

new3 = (
"                       \"CustoUnitario = \" & FSQL(xVUnTrib) & \", \" & _\r\n"
"                       \"DataAtualizacao = GETDATE(), \"\r\n"
)

# ---- R4: CustoUnitario no INSERT usa xVUnTrib em vez de xVUnCom/qtfracao ----
old4 = (
"                       FSQL(xVUnCom / qtfracao) & \",GETDATE())\"\r\n"
)

new4 = (
"                       FSQL(xVUnTrib) & \",GETDATE())\"\r\n"
)

# ---- R5: apos o End If do upsert, atualiza EANEmbalagem e Fracionamento em Produtos ----
old5 = (
"             SQLExecuta sVinSql\r\n"
"          End If\r\n"
"    \r\n"
"           xCfop = XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/CFOP\").Item(0).Text\r\n"
)

new5 = (
"             SQLExecuta sVinSql\r\n"
"          End If\r\n"
"\r\n"
"          '--- Atualiza EANEmbalagem e Fracionamento em Produtos quando ha vinculo ---\r\n"
"          If CodProduto > 0 And Not Vazio(cEAN) Then\r\n"
"             SQLExecuta \"UPDATE Produtos SET EANEmbalagem = '\" & cEAN & \"', \" & _\r\n"
"                        \"Fracionamento = \" & FSQL(qtfracao) & \" WHERE Codigo = \" & CodProduto\r\n"
"          End If\r\n"
"    \r\n"
"           xCfop = XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/CFOP\").Item(0).Text\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3 found:', data.count(old3))
print('r4 found:', data.count(old4))
print('r5 found:', data.count(old5))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3, new3, 1)
data2 = data2.replace(old4, new4, 1)
data2 = data2.replace(old5, new5, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
