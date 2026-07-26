data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Adicionar Dim curBasePISCOFINS no bloco de Dims do Sub
old_dims = b'Dim vCSOSN As String\r\n'
new_dims = (
    b'Dim vCSOSN As String\r\n'
    b'Dim curBasePISCOFINS As Currency\r\n'
)
c = data.count(old_dims)
print(f'1. Dims: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# 2. Substituir secoes 'PIS e 'COFINS
old_piscofins = (
    b"    'PIS\r\n"
    b'    Tb("PISCST") = Right(Format(vPISCST, "@"), 2)\r\n'
    b'    If vPISALIQ = "" Then Tb("PISpPIS") = CDbl(Format(0, "@")) Else Tb("PISpPIS") = CDbl(Format(vPISALIQ, "@"))\r\n'
    b'    vValorPIS = FormatNumber(((CCur(vValorProdutos) * CDbl(vPISALIQ)) / 100), 2)\r\n'
    b'    If vPISALIQ = "" Then Tb("PISvPIS") = CDbl(Format(0, "@")) Else Tb("PISvPIS") = CDbl(Format(vValorPIS, "@"))\r\n'
    b'    \r\n'
    b"    'COFINS\r\n"
    b'    Tb("COFINSCST") = Right(Format(vCOFINSCST, "@"), 2)\r\n'
    b'    If vCOFINSALIQ = "" Then Tb("cofinspcofins") = CDbl(Format(0, "@")) Else Tb("cofinspcofins") = CDbl(Format(vCOFINSALIQ, "@"))\r\n'
    b'    vValorCOFINS = FormatNumber(((CCur(vValorProdutos) * CDbl(vCOFINSALIQ)) / 100), 2)\r\n'
    b'    If vCOFINSALIQ = "" Then Tb("cofinsvcofins") = CDbl(Format(0, "@")) Else Tb("cofinsvcofins") = CDbl(Format(vValorCOFINS, "@"))\r\n'
)
new_piscofins = (
    b"    'PIS e COFINS\r\n"
    b'    If vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5 Then\r\n'
    b"        ' Simples Nacional / MEI: nao destaca PIS/COFINS por item\r\n"
    b'        Tb("PISCST") = Right(Format(vPISCST, "@"), 2)\r\n'
    b'        Tb("PISvBC") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("PISpPIS") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("PISvPIS") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("PISqBCProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("PISvAliqProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("COFINSCST") = Right(Format(vCOFINSCST, "@"), 2)\r\n'
    b'        Tb("COFINSvBC") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("cofinspcofins") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("cofinsvcofins") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("COFINSqBCProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("COFINSvAliqProd") = CDbl(Format(0, "@"))\r\n'
    b'    Else\r\n'
    b"        ' Regime Normal: base = valor liquido - ICMS (Tese do Seculo STF RE 574.706)\r\n"
    b'        curBasePISCOFINS = CCur(vValorProdutos) - vValorICMS\r\n'
    b'        If curBasePISCOFINS < 0 Then curBasePISCOFINS = 0\r\n'
    b'        Tb("PISCST") = Right(Format(vPISCST, "@"), 2)\r\n'
    b'        Tb("PISvBC") = CDbl(Format(curBasePISCOFINS, "0.00"))\r\n'
    b'        If vPISALIQ = "" Then\r\n'
    b'            Tb("PISpPIS") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("PISvPIS") = CDbl(Format(0, "@"))\r\n'
    b'        Else\r\n'
    b'            Tb("PISpPIS") = CDbl(Format(vPISALIQ, "@"))\r\n'
    b'            vValorPIS = CCur(Format(curBasePISCOFINS * CDbl(vPISALIQ) / 100, "0.00"))\r\n'
    b'            Tb("PISvPIS") = CDbl(Format(vValorPIS, "@"))\r\n'
    b'        End If\r\n'
    b'        Tb("PISqBCProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("PISvAliqProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("COFINSCST") = Right(Format(vCOFINSCST, "@"), 2)\r\n'
    b'        Tb("COFINSvBC") = CDbl(Format(curBasePISCOFINS, "0.00"))\r\n'
    b'        If vCOFINSALIQ = "" Then\r\n'
    b'            Tb("cofinspcofins") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("cofinsvcofins") = CDbl(Format(0, "@"))\r\n'
    b'        Else\r\n'
    b'            Tb("cofinspcofins") = CDbl(Format(vCOFINSALIQ, "@"))\r\n'
    b'            vValorCOFINS = CCur(Format(curBasePISCOFINS * CDbl(vCOFINSALIQ) / 100, "0.00"))\r\n'
    b'            Tb("cofinsvcofins") = CDbl(Format(vValorCOFINS, "@"))\r\n'
    b'        End If\r\n'
    b'        Tb("COFINSqBCProd") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("COFINSvAliqProd") = CDbl(Format(0, "@"))\r\n'
    b'    End If\r\n'
)
c = data.count(old_piscofins)
print(f'2. PIS/COFINS: {c}')
if c == 1: data = data.replace(old_piscofins, new_piscofins)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
