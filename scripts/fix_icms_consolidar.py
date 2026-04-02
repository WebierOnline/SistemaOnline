data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Adicionar Dim vCSOSN no bloco de Dims do Sub
old_dims = (
    b'Dim vBCUFDestLDI As Double\r\n'
    b'Dim vICMSUFDestLDI As Double\r\n'
    b'Dim vFCPUFDestLDI  As Double\r\n'
)
new_dims = (
    b'Dim vBCUFDestLDI As Double\r\n'
    b'Dim vICMSUFDestLDI As Double\r\n'
    b'Dim vFCPUFDestLDI  As Double\r\n'
    b'Dim vCSOSN As String\r\n'
)
c = data.count(old_dims)
print(f'1. Dims: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# 2. Substituir secao 'ICMS: mover IPIpIPI para fora e adicionar vBC/vICMS
old_icms = (
    b"    'ICMS\r\n"
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
    b'    If vIPIALIQ = "" Then Tb("IPIpIPI") = CDbl(Format(0, "@")) Else Tb("IPIpIPI") = CDbl(Format(vIPIALIQ, "@"))\r\n'
    b'    If vpRedBC = "" Then Tb("pRedBC") = CDbl(Format(0, "@")) Else Tb("pRedBC") = CDbl(Format(vpRedBC, "@"))\r\n'
)
new_icms = (
    b"    'ICMS\r\n"
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
    b'    If vpRedBC = "" Then Tb("pRedBC") = CDbl(Format(0, "@")) Else Tb("pRedBC") = CDbl(Format(vpRedBC, "@"))\r\n'
    b'    If vTipoCRT = 1 And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
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
c = data.count(old_icms)
print(f'2. Secao ICMS: {c}')
if c == 1: data = data.replace(old_icms, new_icms)

# 3. Adicionar IPIpIPI de volta na secao IPI (onde pertence)
old_ipi = (
    b"    'IPI\r\n"
    b'    vPorcIPI = vIPIALIQ\r\n'
)
new_ipi = (
    b"    'IPI\r\n"
    b'    If vIPIALIQ = "" Then Tb("IPIpIPI") = CDbl(Format(0, "@")) Else Tb("IPIpIPI") = CDbl(Format(vIPIALIQ, "@"))\r\n'
    b'    vPorcIPI = vIPIALIQ\r\n'
)
c = data.count(old_ipi)
print(f'3. IPIpIPI para secao IPI: {c}')
if c == 1: data = data.replace(old_ipi, new_ipi)

# 4. Remover secao ' vBC e vICMS que ficou entre Valores do item
old_vbc_section = (
    b'    Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'    \' vBC e vICMS\r\n'
    b'    If vTipoCRT = 1 And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b"        ' Simples Nacional: zera vBC/vICMS exceto CSOSN 101 e 201 (permite credito)\r\n"
    b'        Dim vCSOSN As String\r\n'
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
    b"        ' Regime Normal: aplica reducao de BC se houver\r\n"
    b'        If vpRedBC <> "" And CDbl(vpRedBC) > 0 Then\r\n'
    b'            Tb("vBC") = CDbl(Format(CCur(txtSubTotal.Text) * (1 - CDbl(vpRedBC) / 100), "0.00"))\r\n'
    b'        Else\r\n'
    b'            Tb("vBC") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'        End If\r\n'
    b'        vPorcICMS = CDbl(IIf(vICMSAliq = "", 0, vICMSAliq))\r\n'
    b'        vValorICMS = CCur(Format((CDbl(Tb("vBC")) * vPorcICMS) / 100, "0.00"))\r\n'
    b'        If vICMSAliq = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
    b'    End If\r\n'
    b'    \r\n'
    b'    Tb("referencia") = Format(0, "@")\r\n'
)
new_vbc_section = (
    b'    Tb("ValorTotalBruto") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'    \r\n'
    b'    Tb("referencia") = Format(0, "@")\r\n'
)
c = data.count(old_vbc_section)
print(f'4. Remover secao vBC duplicada: {c}')
if c == 1: data = data.replace(old_vbc_section, new_vbc_section)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
