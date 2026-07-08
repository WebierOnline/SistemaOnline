data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Adicionar Dim curBaseICMS no bloco de Dims do Sub
old_dims = b'Dim curBasePISCOFINS As Currency\r\n'
new_dims = (
    b'Dim curBasePISCOFINS As Currency\r\n'
    b'Dim curBaseICMS As Currency\r\n'
)
c = data.count(old_dims)
print(f'1. Dims: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# 2. Substituir secao 'ICMS completa
old_icms = (
    b"    'ICMS\r\n"
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
    b'    If vpRedBC = "" Then Tb("pRedBC") = CDbl(Format(0, "@")) Else Tb("pRedBC") = CDbl(Format(vpRedBC, "@"))\r\n'
    b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b"        ' Simples Nacional: zera vBC/vICMS exceto CSOSN 101 e 201 (permite credito)\r\n"
    b'        vCSOSN = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'        If vCSOSN = "101" Or vCSOSN = "201" Then\r\n'
    b'            vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))\r\n'
    b'            Tb("vBC") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'            vValorICMS = CCur(Format((CCur(txtSubTotal.Text) * vPorcICMS) / 100, "0.00"))\r\n'
    b'            Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
    b'        Else\r\n'
    b'            Tb("vBC") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("vICMS") = CDbl(Format(0, "@"))\r\n'
    b'        End If\r\n'
    b'    Else\r\n'
    b"        ' Regime Normal ou devolucao: aplica reducao de BC se houver\r\n"
    b'        If vpRedBC <> "" And CDbl(vpRedBC) > 0 Then\r\n'
    b'            Tb("vBC") = CDbl(Format(CCur(txtSubTotal.Text) * (1 - CDbl(vpRedBC) / 100), "0.00"))\r\n'
    b'        Else\r\n'
    b'            Tb("vBC") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'        End If\r\n'
    b'        vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))\r\n'
    b'        vValorICMS = CCur(Format((CDbl(Tb("vBC")) * vPorcICMS) / 100, "0.00"))\r\n'
    b'        If vICMSAliq = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
    b'    End If\r\n'
)
new_icms = (
    b"    'ICMS\r\n"
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
    b'    If vpRedBC = "" Then Tb("pRedBC") = CDbl(Format(0, "@")) Else Tb("pRedBC") = CDbl(Format(vpRedBC, "@"))\r\n'
    b"    ' Calculo antecipado do IPI para usar na base do ICMS (consumidor final)\r\n"
    b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b'        vValorIPI = 0\r\n'
    b'    Else\r\n'
    b'        If vIPIALIQ = "" Then\r\n'
    b'            vValorIPI = 0\r\n'
    b'        Else\r\n'
    b'            vValorIPI = CCur(Format(CCur(vValorProdutos) * CDbl(vIPIALIQ) / 100, "0.00"))\r\n'
    b'        End If\r\n'
    b'    End If\r\n'
    b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b"        ' Simples Nacional venda normal: zera vBC/vICMS exceto CSOSN 101 e 201\r\n"
    b'        vCSOSN = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'        If vCSOSN = "101" Or vCSOSN = "201" Then\r\n'
    b'            vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))\r\n'
    b'            curBaseICMS = CCur(txtSubTotal.Text)\r\n'
    b'            If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI\r\n'
    b'            Tb("vBC") = CDbl(Format(curBaseICMS, "0.00"))\r\n'
    b'            vValorICMS = CCur(Format(curBaseICMS * vPorcICMS / 100, "0.00"))\r\n'
    b'            Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
    b'        Else\r\n'
    b'            Tb("vBC") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("vICMS") = CDbl(Format(0, "@"))\r\n'
    b'        End If\r\n'
    b'    Else\r\n'
    b"        ' Regime Normal ou devolucao: aplica reducao de BC + IPI se consumidor final\r\n"
    b'        If vpRedBC <> "" And CDbl(vpRedBC) > 0 Then\r\n'
    b'            curBaseICMS = CCur(txtSubTotal.Text) * (1 - CDbl(vpRedBC) / 100)\r\n'
    b'        Else\r\n'
    b'            curBaseICMS = CCur(txtSubTotal.Text)\r\n'
    b'        End If\r\n'
    b'        If Left(cboConsumidorFinal.Text, 1) = "1" Then curBaseICMS = curBaseICMS + vValorIPI\r\n'
    b'        Tb("vBC") = CDbl(Format(curBaseICMS, "0.00"))\r\n'
    b'        vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))\r\n'
    b'        vValorICMS = CCur(Format(curBaseICMS * vPorcICMS / 100, "0.00"))\r\n'
    b'        If vICMSAliq = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
    b'    End If\r\n'
)
c = data.count(old_icms)
print(f'2. Secao ICMS: {c}')
if c == 1: data = data.replace(old_icms, new_icms)

# 3. Substituir secao 'IPI
old_ipi = (
    b"    'IPI\r\n"
    b'    If vIPIALIQ = "" Then Tb("IPIpIPI") = CDbl(Format(0, "@")) Else Tb("IPIpIPI") = CDbl(Format(vIPIALIQ, "@"))\r\n'
    b'    vPorcIPI = vIPIALIQ\r\n'
    b'    vValorIPI = FormatNumber(((CCur(vValorProdutos) * CDbl(vPorcIPI)) / 100), 2)\r\n'
    b'    If vIPIALIQ = "" Then Tb("IPIvIPI") = CDbl(Format(0, "@")) Else Tb("IPIvIPI") = CDbl(Format(vValorIPI, "@"))\r\n'
    b'    Tb("IPICST") = Format(vIPICST, "@")\r\n'
    b'    If vIPICST = "99" Or vIPICST = "53" Or vIPICST = "52" Or vIPICST = "50" Then\r\n'
    b'        Tb("IPIcEnq") = "999"\r\n'
    b'    Else\r\n'
    b'        Tb("IPIcEnq") = ""\r\n'
    b'    End If\r\n'
)
new_ipi = (
    b"    'IPI\r\n"
    b'    Tb("IPICST") = Format(vIPICST, "@")\r\n'
    b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b"        ' Simples Nacional venda normal: zera IPI\r\n"
    b'        Tb("IPIcEnq") = "999"\r\n'
    b'        Tb("IPIvBC") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("IPIpIPI") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("IPIvIPI") = CDbl(Format(0, "@"))\r\n'
    b'    Else\r\n'
    b"        ' Regime Normal ou devolucao: calcula IPI (vValorIPI ja calculado na secao ICMS)\r\n"
    b'        If vIPICST = "99" Or vIPICST = "53" Or vIPICST = "52" Or vIPICST = "50" Then\r\n'
    b'            Tb("IPIcEnq") = "999"\r\n'
    b'        Else\r\n'
    b'            Tb("IPIcEnq") = ""\r\n'
    b'        End If\r\n'
    b'        Tb("IPIvBC") = CDbl(Format(vValorProdutos, "0.00"))\r\n'
    b'        If vIPIALIQ = "" Then\r\n'
    b'            Tb("IPIpIPI") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("IPIvIPI") = CDbl(Format(0, "@"))\r\n'
    b'        Else\r\n'
    b'            Tb("IPIpIPI") = CDbl(Format(vIPIALIQ, "@"))\r\n'
    b'            Tb("IPIvIPI") = CDbl(Format(vValorIPI, "@"))\r\n'
    b'        End If\r\n'
    b'    End If\r\n'
)
c = data.count(old_ipi)
print(f'3. Secao IPI: {c}')
if c == 1: data = data.replace(old_ipi, new_ipi)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
