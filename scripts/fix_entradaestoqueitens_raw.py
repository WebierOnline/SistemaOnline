path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

old = (
"           Compras_Itens_XML!QuantidadeComercial = (CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qCom\").Item(0).Text, \".\", \",\", UM_A_UM)) * qtfracao)          'qCom: Quantidade comercial <qCom>\r\n"
"           Compras_Itens_XML!ValorUnitarioComercializacao = (CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/vUnCom\").Item(0).Text, \".\", \",\", UM_A_UM)) / qtfracao) 'vUnCom: Valor unitario de comercializacao <vUnCom>\r\n"
)

new = (
"           Compras_Itens_XML!QuantidadeComercial = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/qCom\").Item(0).Text, \".\", \",\", UM_A_UM))          'qCom: Quantidade comercial <qCom>\r\n"
"           Compras_Itens_XML!ValorUnitarioComercializacao = CDbl(Substitui(XMLDOC.selectNodes(\"nfeProc/NFe/infNFe/det\").Item(i).selectNodes(\"prod/vUnCom\").Item(0).Text, \".\", \",\", UM_A_UM)) 'vUnCom: Valor unitario de comercializacao <vUnCom>\r\n"
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
