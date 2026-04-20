path = 'C:/projeto/OnlineCommerce/Forms/frmVinculoProdutoXML.frm'
data = open(path, 'rb').read().decode('windows-1252')

# ---- R1: adiciona vUnTrib ao tItemXML ----
old1 = (
"   uCom       As String\r\n"
"   vUnCom     As Double\r\n"
)

new1 = (
"   uCom       As String\r\n"
"   vUnCom     As Double\r\n"
"   vUnTrib    As Double\r\n"
)

# ---- R2: adiciona ValorUnitarioTributario na query ----
old2 = (
"              \"       ValorUnitarioComercializacao AS vUnCom, \" & _\r\n"
"              \"       ISNULL(NCM,'') AS NCM, ISNULL(CEST,'') AS CEST, \" & _\r\n"
)

new2 = (
"              \"       ValorUnitarioComercializacao AS vUnCom, \" & _\r\n"
"              \"       ISNULL(ValorUnitarioTributario,0) AS vUnTrib, \" & _\r\n"
"              \"       ISNULL(NCM,'') AS NCM, ISNULL(CEST,'') AS CEST, \" & _\r\n"
)

# ---- R3a: inicializa .vUnTrib = 0 junto com .vUnCom = 0 ----
old3a = (
"         .vUnCom = 0\r\n"
"         .ICMSAliq = 0: .pRedBC = 0: .modBC = 3\r\n"
)

new3a = (
"         .vUnCom = 0\r\n"
"         .vUnTrib = 0\r\n"
"         .ICMSAliq = 0: .pRedBC = 0: .modBC = 3\r\n"
)

# ---- R3b: popula .vUnTrib do recordset ----
old3b = (
"         .vUnCom = CDbl(rs!vUnCom)\r\n"
"         .ICMSAliq = CDbl(rs!ICMSAliq)\r\n"
)

new3b = (
"         .vUnCom = CDbl(rs!vUnCom)\r\n"
"         .vUnTrib = CDbl(rs!vUnTrib)\r\n"
"         .ICMSAliq = CDbl(rs!ICMSAliq)\r\n"
)

# ---- R4: apos salvar VinculoXMLProduto, atualiza Produtos.EANEmbalagem + Fracionamento ----
old4 = (
"   Dim msgErro As String\r\n"
"   msgErro = SQLExecuta(sSQL)\r\n"
"   If Not Vazio(msgErro) Then\r\n"
"      MsgBox \"Erro ao salvar v\" & Chr(237) & \"nculo: \" & msgErro, vbCritical\r\n"
"      Exit Function\r\n"
"   End If\r\n"
"   \r\n"
"   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens\r\n"
)

new4 = (
"   Dim msgErro As String\r\n"
"   msgErro = SQLExecuta(sSQL)\r\n"
"   If Not Vazio(msgErro) Then\r\n"
"      MsgBox \"Erro ao salvar v\" & Chr(237) & \"nculo: \" & msgErro, vbCritical\r\n"
"      Exit Function\r\n"
"   End If\r\n"
"   \r\n"
"   'Atualiza EANEmbalagem e Fracionamento em Produtos\r\n"
"   SQLExecuta \"UPDATE Produtos SET EANEmbalagem = '\" & sEANEmb & \"', \" & _\r\n"
"              \"Fracionamento = \" & FSQL(frac) & \" WHERE Codigo = \" & IDProdSel\r\n"
"   \r\n"
"   'Atualiza CodigoProduto + quantidades em EntradaEstoqueItens\r\n"
)

print('r1 found:', data.count(old1))
print('r2 found:', data.count(old2))
print('r3a found:', data.count(old3a))
print('r3b found:', data.count(old3b))
print('r4 found:', data.count(old4))

data2 = data
data2 = data2.replace(old1, new1, 1)
data2 = data2.replace(old2, new2, 1)
data2 = data2.replace(old3a, new3a, 1)
data2 = data2.replace(old3b, new3b, 1)
data2 = data2.replace(old4, new4, 1)

print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
