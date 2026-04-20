path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: adiciona sVinUnidMedida na declaracao e leitura junto com sVinEANProduto ----
old1 = (
"          Dim sVinxProd As String, sVinuCom As String, sVinEANProduto As String\r\n"
"          Dim sVinSql As String\r\n"
)
new1 = (
"          Dim sVinxProd As String, sVinuCom As String, sVinEANProduto As String, sVinUnidMedida As String\r\n"
"          Dim sVinSql As String\r\n"
)

old2 = (
"          If CodProduto > 0 Then\r\n"
"             sVinEANProduto = SQLExecutaRetorno(\"SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = \" & CodProduto, \"r\", \"\")\r\n"
"          End If\r\n"
"          If bVinculoExiste Then\r\n"
)
new2 = (
"          sVinUnidMedida = \"\"\r\n"
"          If CodProduto > 0 Then\r\n"
"             sVinEANProduto = SQLExecutaRetorno(\"SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = \" & CodProduto, \"r\", \"\")\r\n"
"             sVinUnidMedida = SQLExecutaRetorno(\"SELECT ISNULL(UNID_MEDIDA,'') r FROM Produtos WHERE Codigo = \" & CodProduto, \"r\", \"\")\r\n"
"          End If\r\n"
"          If bVinculoExiste Then\r\n"
)

# ---- R3: adiciona QuantidadeComercial e ValorUnitarioComercializacao no UPDATE (parte 2) ----
old3 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"PISCST = '\" & PISCST & \"', \" & _\r\n"
"                       \"PISpPIS = \" & FSQL(PISpPIS) & \", \" & _\r\n"
"                       \"COFINSCST = '\" & COFINSCST & \"', \" & _\r\n"
"                       \"COFINSpCOFINS = \" & FSQL(COFINSpCOFINS) & \", \" & _\r\n"
"                       \"CEST = '\" & sVinCEST & \"', \" & _\r\n"
"                       \"CustoUnitario = \" & FSQL(xVUnTrib) & \", \" & _\r\n"
"                       \"DataAtualizacao = GETDATE(), \"\r\n"
)
new3 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"PISCST = '\" & PISCST & \"', \" & _\r\n"
"                       \"PISpPIS = \" & FSQL(PISpPIS) & \", \" & _\r\n"
"                       \"COFINSCST = '\" & COFINSCST & \"', \" & _\r\n"
"                       \"COFINSpCOFINS = \" & FSQL(COFINSpCOFINS) & \", \" & _\r\n"
"                       \"CEST = '\" & sVinCEST & \"', \" & _\r\n"
"                       \"QuantidadeComercial = \" & FSQL(xQCom) & \", \" & _\r\n"
"                       \"ValorUnitarioComercializacao = \" & FSQL(xVUnCom) & \", \" & _\r\n"
"                       \"CustoUnitario = \" & FSQL(xVUnTrib) & \", \" & _\r\n"
"                       \"DataAtualizacao = GETDATE(), \"\r\n"
)

# ---- R4: adiciona UNID_MEDIDA no UPDATE (parte 3, CASE WHEN) ----
old4 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"IDProduto = CASE WHEN IDProduto = 0 THEN \" & CodProduto & \" ELSE IDProduto END, \" & _\r\n"
"                       \"EANProduto = CASE WHEN IDProduto = 0 THEN '\" & sVinEANProduto & \"' ELSE EANProduto END, \" & _\r\n"
"                       \"Fracionamento = CASE WHEN IDProduto = 0 THEN \" & FSQL(qtfracao) & \" ELSE Fracionamento END \" & _\r\n"
"                       \"WHERE IDFornecedor = \" & Codigo_fornecedor & \" AND cProd = '\" & Replace(xCProd, \"'\", \"''\") & \"'\"\r\n"
)
new4 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"IDProduto = CASE WHEN IDProduto = 0 THEN \" & CodProduto & \" ELSE IDProduto END, \" & _\r\n"
"                       \"EANProduto = CASE WHEN IDProduto = 0 THEN '\" & sVinEANProduto & \"' ELSE EANProduto END, \" & _\r\n"
"                       \"UNID_MEDIDA = CASE WHEN IDProduto = 0 THEN '\" & sVinUnidMedida & \"' ELSE UNID_MEDIDA END, \" & _\r\n"
"                       \"Fracionamento = CASE WHEN IDProduto = 0 THEN \" & FSQL(qtfracao) & \" ELSE Fracionamento END \" & _\r\n"
"                       \"WHERE IDFornecedor = \" & Codigo_fornecedor & \" AND cProd = '\" & Replace(xCProd, \"'\", \"''\") & \"'\"\r\n"
)

# ---- R5: adiciona 3 campos no INSERT (campo list) ----
old5 = (
"             sVinSql = \"INSERT INTO VinculoXMLProduto (IDFornecedor,cProd,EANEmbalagem,xProd,uCom,\" & _\r\n"
"                       \"NCM,CFOP,CST,pICMS,IPICST,IPIpIPI,PISCST,PISpPIS,COFINSCST,COFINSpCOFINS,CEST,\" & _\r\n"
"                       \"IDProduto,EANProduto,Fracionamento,CustoUnitario,DataAtualizacao) VALUES (\" & _\r\n"
"                       Codigo_fornecedor & \",'\" & Replace(xCProd, \"'\", \"''\") & \"','\" & cEAN & \"',\"\r\n"
)
new5 = (
"             sVinSql = \"INSERT INTO VinculoXMLProduto (IDFornecedor,cProd,EANEmbalagem,xProd,uCom,\" & _\r\n"
"                       \"QuantidadeComercial,ValorUnitarioComercializacao,UNID_MEDIDA,\" & _\r\n"
"                       \"NCM,CFOP,CST,pICMS,IPICST,IPIpIPI,PISCST,PISpPIS,COFINSCST,COFINSpCOFINS,CEST,\" & _\r\n"
"                       \"IDProduto,EANProduto,Fracionamento,CustoUnitario,DataAtualizacao) VALUES (\" & _\r\n"
"                       Codigo_fornecedor & \",'\" & Replace(xCProd, \"'\", \"''\") & \"','\" & cEAN & \"',\"\r\n"
)

# ---- R6: adiciona valores dos 3 campos no INSERT (values, parte 2) ----
old6 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"'\" & Replace(sVinxProd, \"'\", \"''\") & \"','\" & sVinuCom & \"',\" & _\r\n"
"                       \"'\" & sVinNCM & \"',\" & sVinCFOP & \",'\" & Replace(csticms, \"'\", \"''\") & \"',\" & _\r\n"
"                       FSQL(ICMSAliquota) & \",'\" & cstipi & \"',\" & FSQL(IPIpIPI) & \",'\" & PISCST & \"',\"\r\n"
)
new6 = (
"             sVinSql = sVinSql & _\r\n"
"                       \"'\" & Replace(sVinxProd, \"'\", \"''\") & \"','\" & sVinuCom & \"',\" & _\r\n"
"                       FSQL(xQCom) & \",\" & FSQL(xVUnCom) & \",'\" & sVinUnidMedida & \"',\" & _\r\n"
"                       \"'\" & sVinNCM & \"',\" & sVinCFOP & \",'\" & Replace(csticms, \"'\", \"''\") & \"',\" & _\r\n"
"                       FSQL(ICMSAliquota) & \",'\" & cstipi & \"',\" & FSQL(IPIpIPI) & \",'\" & PISCST & \"',\"\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3 found:', data.count(old3))
print('r4 found:', data.count(old4))
print('r5 found:', data.count(old5))
print('r6 found:', data.count(old6))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3, new3, 1)
data2 = data2.replace(old4, new4, 1)
data2 = data2.replace(old5, new5, 1)
data2 = data2.replace(old6, new6, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
