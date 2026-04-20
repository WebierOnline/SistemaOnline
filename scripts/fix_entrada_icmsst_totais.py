data = open('OnlineCommerce/Forms/Entrada_Estoque.Frm', 'rb').read()

# ============================================================
# 1. Declarar acumuladores antes do For i = 0 To qtdProd - 1
# ============================================================
old1 = (
    b"       End If\r\n"
    b"\r\n"
    b"       For i = 0 To qtdProd - 1 'Varrendo todos os itens\r\n"
)

new1 = (
    b"       End If\r\n"
    b"\r\n"
    b"       ' Acumuladores totais ICMS60 (ST retido anteriormente) para cabecalho\r\n"
    b"       Dim dTotalBCSTRet As Double\r\n"
    b"       Dim dTotalICMSSTRet As Double\r\n"
    b"       Dim dTotalICMSSubstituto As Double\r\n"
    b"       dTotalBCSTRet = 0: dTotalICMSSTRet = 0: dTotalICMSSubstituto = 0\r\n"
    b"\r\n"
    b"       For i = 0 To qtdProd - 1 'Varrendo todos os itens\r\n"
)

c = data.count(old1)
print(f'1. Acumuladores antes do For: count={c}')
if c == 1:
    data = data.replace(old1, new1)

# ============================================================
# 2. Acumular dentro do bloco ICMS60, apos o ultimo End If
#    (antes de Case "ICMS70")
# ============================================================
old2 = (
    b"                 If oICMS.selectNodes(\"ICMS60/vFCPSTRet\").Length > 0 Then\r\n"
    b"                    ICMSvFCPSTRet = CDbl(Substitui(oICMS.selectNodes(\"ICMS60/vFCPSTRet\").item(0).Text, \".\", \",\", UM_A_UM))\r\n"
    b"                 End If\r\n"
    b"\r\n"
    b"              Case \"ICMS70\"\r\n"
)

new2 = (
    b"                 If oICMS.selectNodes(\"ICMS60/vFCPSTRet\").Length > 0 Then\r\n"
    b"                    ICMSvFCPSTRet = CDbl(Substitui(oICMS.selectNodes(\"ICMS60/vFCPSTRet\").item(0).Text, \".\", \",\", UM_A_UM))\r\n"
    b"                 End If\r\n"
    b"                 ' Acumula para total no cabecalho\r\n"
    b"                 dTotalBCSTRet = dTotalBCSTRet + ICMSvBCSTRet\r\n"
    b"                 dTotalICMSSTRet = dTotalICMSSTRet + ICMSvICMSSTRet\r\n"
    b"                 dTotalICMSSubstituto = dTotalICMSSubstituto + ICMSvICMSSubstituto\r\n"
    b"\r\n"
    b"              Case \"ICMS70\"\r\n"
)

c = data.count(old2)
print(f'2. Acumular no ICMS60: count={c}')
if c == 1:
    data = data.replace(old2, new2)

# ============================================================
# 3. Apos Next i, gravar os totais com UPDATE
# ============================================================
old3 = (
    b"    Next i\r\n"
    b"    '------------- FIM DA IMPORTAaaO DOS ITENS ------------------\r\n"
)

new3 = (
    b"    Next i\r\n"
    b"    ' Gravar totais ICMS60 no cabecalho da entrada\r\n"
    b"    dbData.Execute \"UPDATE EntradaEstoque SET \" & _\r\n"
    b"        \"vBCSTRetTotal = \" & Replace(Format(dTotalBCSTRet, \"0.00\"), \",\", \".\") & \", \" & _\r\n"
    b"        \"vICMSSTRetTotal = \" & Replace(Format(dTotalICMSSTRet, \"0.00\"), \",\", \".\") & \", \" & _\r\n"
    b"        \"vICMSSubstitutoTotal = \" & Replace(Format(dTotalICMSSubstituto, \"0.00\"), \",\", \".\") & _\r\n"
    b"        \" WHERE CodigoNota = \" & Compra\r\n"
    b"    '------------- FIM DA IMPORTAaaO DOS ITENS ------------------\r\n"
)

c = data.count(old3)
print(f'3. UPDATE apos Next i: count={c}')
if c == 1:
    data = data.replace(old3, new3)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/Entrada_Estoque.Frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
