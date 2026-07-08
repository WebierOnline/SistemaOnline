path = 'C:/projeto/OnlineCommerce/Forms/frmVinculoProdutoXML.frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: adiciona qCom ao tipo tItemXML ----
old1 = (
"   vUnCom     As Double\r\n"
"   vUnTrib    As Double\r\n"
)
new1 = (
"   vUnCom     As Double\r\n"
"   vUnTrib    As Double\r\n"
"   qCom       As Double\r\n"
)

# ---- R2: adiciona QuantidadeComercial AS qCom no SELECT ----
old2 = (
"              \"       ValorUnitarioComercializacao AS vUnCom, \" & _\r\n"
"              \"       ISNULL(ValorUnitarioTributario,0) AS vUnTrib, \" & _\r\n"
)
new2 = (
"              \"       ValorUnitarioComercializacao AS vUnCom, \" & _\r\n"
"              \"       ISNULL(QuantidadeComercial,0) AS qCom, \" & _\r\n"
"              \"       ISNULL(ValorUnitarioTributario,0) AS vUnTrib, \" & _\r\n"
)

# ---- R3a: inicializa .qCom = 0 ----
old3a = (
"         .vUnCom = 0\r\n"
"         .vUnTrib = 0\r\n"
)
new3a = (
"         .vUnCom = 0\r\n"
"         .vUnTrib = 0\r\n"
"         .qCom = 0\r\n"
)

# ---- R3b: popula .qCom do recordset ----
old3b = (
"         .vUnCom = CDbl(rs!vUnCom)\r\n"
"         .vUnTrib = CDbl(rs!vUnTrib)\r\n"
)
new3b = (
"         .vUnCom = CDbl(rs!vUnCom)\r\n"
"         .vUnTrib = CDbl(rs!vUnTrib)\r\n"
"         .qCom = CDbl(rs!qCom)\r\n"
)

# ---- R4: custoUnit usa vUnTrib em vez de vUnCom / frac ----
old4 = (
"   Dim custoUnit As Double\r\n"
"   custoUnit = IIf(frac > 0, Item.vUnCom / frac, Item.vUnCom)\r\n"
)
new4 = (
"   Dim custoUnit As Double\r\n"
"   custoUnit = Item.vUnTrib\r\n"
)

# ---- R5a: adiciona sUnidMedida apos sEANProd ----
old5a = (
"   'EAN do produto interno\r\n"
"   Dim sEANProd As String\r\n"
"   sEANProd = Trim(SQLExecutaRetorno(\"SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = \" & IDProdSel, \"r\", \"\"))\r\n"
"\r\n"
"   'IDFornecedor\r\n"
)
new5a = (
"   'EAN do produto interno\r\n"
"   Dim sEANProd As String\r\n"
"   sEANProd = Trim(SQLExecutaRetorno(\"SELECT ISNULL(EAN,'') r FROM Produtos WHERE Codigo = \" & IDProdSel, \"r\", \"\"))\r\n"
"\r\n"
"   'Unidade de medida interna\r\n"
"   Dim sUnidMedida As String\r\n"
"   sUnidMedida = Trim(SQLExecutaRetorno(\"SELECT ISNULL(UNID_MEDIDA,'') r FROM Produtos WHERE Codigo = \" & IDProdSel, \"r\", \"\"))\r\n"
"\r\n"
"   'IDFornecedor\r\n"
)

# ---- R5b: expande INSERT com os 3 novos campos ----
old5b = (
"      sSQL = \"INSERT INTO VinculoXMLProduto \" & _\r\n"
"             \"(IDFornecedor, cProd, EANEmbalagem, xProd, uCom, IDProduto, EANProduto, Fracionamento, CustoUnitario, DataAtualizacao) \" & _\r\n"
"             \"VALUES (\" & _\r\n"
"             IDForn & \", \" & _\r\n"
"             \"'\" & Replace(Item.cProd, \"'\", \"''\") & \"', \" & _\r\n"
"             \"'\" & sEANEmb & \"', \" & _\r\n"
"             \"'\" & Replace(Item.Nome, \"'\", \"''\") & \"', \" & _\r\n"
"             \"'\" & Item.uCom & \"', \" & _\r\n"
"             IDProdSel & \", \" & _\r\n"
"             \"'\" & sEANProd & \"', \" & _\r\n"
"             FSQL(frac) & \", \" & _\r\n"
"             FSQL(custoUnit) & \", GETDATE())\"\r\n"
)
new5b = (
"      sSQL = \"INSERT INTO VinculoXMLProduto \" & _\r\n"
"             \"(IDFornecedor, cProd, EANEmbalagem, xProd, uCom, QuantidadeComercial, ValorUnitarioComercializacao, \" & _\r\n"
"             \"IDProduto, EANProduto, UNID_MEDIDA, Fracionamento, CustoUnitario, DataAtualizacao) \" & _\r\n"
"             \"VALUES (\" & _\r\n"
"             IDForn & \", \" & _\r\n"
"             \"'\" & Replace(Item.cProd, \"'\", \"''\") & \"', \" & _\r\n"
"             \"'\" & sEANEmb & \"', \" & _\r\n"
"             \"'\" & Replace(Item.Nome, \"'\", \"''\") & \"', \" & _\r\n"
"             \"'\" & Item.uCom & \"', \" & _\r\n"
"             FSQL(Item.qCom) & \", \" & _\r\n"
"             FSQL(Item.vUnCom) & \", \" & _\r\n"
"             IDProdSel & \", \" & _\r\n"
"             \"'\" & sEANProd & \"', \" & _\r\n"
"             \"'\" & sUnidMedida & \"', \" & _\r\n"
"             FSQL(frac) & \", \" & _\r\n"
"             FSQL(custoUnit) & \", GETDATE())\"\r\n"
)

# ---- R5c: expande UPDATE com os 3 novos campos ----
old5c = (
"      sSQL = \"UPDATE VinculoXMLProduto SET \" & _\r\n"
"             \"IDProduto = \" & IDProdSel & \", \" & _\r\n"
"             \"EANProduto = '\" & sEANProd & \"', \" & _\r\n"
"             \"EANEmbalagem = '\" & sEANEmb & \"', \" & _\r\n"
"             \"Fracionamento = \" & FSQL(frac) & \", \" & _\r\n"
"             \"CustoUnitario = \" & FSQL(custoUnit) & \", \" & _\r\n"
"             \"DataAtualizacao = GETDATE() \" & _\r\n"
"             \"WHERE IDFornecedor = \" & IDForn & _\r\n"
"             \"  AND cProd = '\" & Replace(Item.cProd, \"'\", \"''\") & \"'\"\r\n"
)
new5c = (
"      sSQL = \"UPDATE VinculoXMLProduto SET \" & _\r\n"
"             \"EANEmbalagem = '\" & sEANEmb & \"', \" & _\r\n"
"             \"xProd = '\" & Replace(Item.Nome, \"'\", \"''\") & \"', \" & _\r\n"
"             \"uCom = '\" & Item.uCom & \"', \" & _\r\n"
"             \"QuantidadeComercial = \" & FSQL(Item.qCom) & \", \" & _\r\n"
"             \"ValorUnitarioComercializacao = \" & FSQL(Item.vUnCom) & \", \" & _\r\n"
"             \"IDProduto = \" & IDProdSel & \", \" & _\r\n"
"             \"EANProduto = '\" & sEANProd & \"', \" & _\r\n"
"             \"UNID_MEDIDA = '\" & sUnidMedida & \"', \" & _\r\n"
"             \"Fracionamento = \" & FSQL(frac) & \", \" & _\r\n"
"             \"CustoUnitario = \" & FSQL(custoUnit) & \", \" & _\r\n"
"             \"DataAtualizacao = GETDATE() \" & _\r\n"
"             \"WHERE IDFornecedor = \" & IDForn & _\r\n"
"             \"  AND cProd = '\" & Replace(Item.cProd, \"'\", \"''\") & \"'\"\r\n"
)

# ---- R6: EntradaEstoqueItens - apenas CodigoProduto, sem alterar dados brutos ----
old6 = (
"   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens\r\n"
"   sSQL = \"UPDATE EntradaEstoqueItens SET \" & _\r\n"
"          \"CodigoProduto = \" & IDProdSel & \", \" & _\r\n"
"          \"QuantidadeComercial = QuantidadeComercial * \" & FSQL(frac) & \", \" & _\r\n"
"          \"ValorUnitarioComercializacao = ValorUnitarioComercializacao / \" & FSQL(frac) & \" \" & _\r\n"
"          \"WHERE CodigoNota = \" & NumeroEntrada & _\r\n"
"          \"  AND Referencia = '\" & Replace(Item.cProd, \"'\", \"''\") & \"' \" & _\r\n"
"          \"  AND (CodigoProduto = 0 OR CodigoProduto IS NULL)\"\r\n"
)
new6 = (
"   'Vincula CodigoProduto em EntradaEstoqueItens (dados brutos da XML preservados)\r\n"
"   sSQL = \"UPDATE EntradaEstoqueItens SET \" & _\r\n"
"          \"CodigoProduto = \" & IDProdSel & \" \" & _\r\n"
"          \"WHERE CodigoNota = \" & NumeroEntrada & _\r\n"
"          \"  AND Referencia = '\" & Replace(Item.cProd, \"'\", \"''\") & \"' \" & _\r\n"
"          \"  AND (CodigoProduto = 0 OR CodigoProduto IS NULL)\"\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3a found:', data.count(old3a))
print('r3b found:', data.count(old3b))
print('r4 found:', data.count(old4))
print('r5a found:', data.count(old5a))
print('r5b found:', data.count(old5b))
print('r5c found:', data.count(old5c))
print('r6 found:', data.count(old6))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3a, new3a, 1)
data2 = data2.replace(old3b, new3b, 1)
data2 = data2.replace(old4, new4, 1)
data2 = data2.replace(old5a, new5a, 1)
data2 = data2.replace(old5b, new5b, 1)
data2 = data2.replace(old5c, new5c, 1)
data2 = data2.replace(old6, new6, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
