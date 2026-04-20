path = 'C:/projeto/OnlineCommerce/Forms/frmVinculoProdutoXML.frm'
data = open(path, 'rb').read().decode('windows-1252')

old4 = (
"   Dim msgErro As String\r\n"
"   msgErro = SQLExecuta(sSQL)\r\n"
"   If Not Vazio(msgErro) Then\r\n"
"      MsgBox \"Erro ao salvar v\" & Chr(237) & \"nculo: \" & msgErro, vbCritical\r\n"
"      Exit Function\r\n"
"   End If\r\n"
"\r\n"
"   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens\r\n"
)

new4 = (
"   Dim msgErro As String\r\n"
"   msgErro = SQLExecuta(sSQL)\r\n"
"   If Not Vazio(msgErro) Then\r\n"
"      MsgBox \"Erro ao salvar v\" & Chr(237) & \"nculo: \" & msgErro, vbCritical\r\n"
"      Exit Function\r\n"
"   End If\r\n"
"\r\n"
"   'Atualiza EANEmbalagem e Fracionamento em Produtos\r\n"
"   SQLExecuta \"UPDATE Produtos SET EANEmbalagem = '\" & sEANEmb & \"', \" & _\r\n"
"              \"Fracionamento = \" & FSQL(frac) & \" WHERE Codigo = \" & IDProdSel\r\n"
"\r\n"
"   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens\r\n"
)

print('r4 found:', data.count(old4))
data2 = data.replace(old4, new4, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
