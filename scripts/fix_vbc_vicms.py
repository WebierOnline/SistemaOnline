data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'    Tb("vBC") = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'    \r\n'
    b'    \r\n'
    b'    \r\n'
    b'    \r\n'
    b'    \r\n'
    b'    \r\n'
    b'    Tb("referencia") = Format(0, "@")\r\n'
    b'    \'calculo do icms de cada produto\r\n'
    b'    \'vValorProdutos = txtSubTotal.Text\r\n'
    b'    \'vPorcICMS = vICMSAliq\r\n'
    b'    \'vValorICMS = Format(((CCur(vValorProdutos) * CDbl(vPorcICMS)) / 100), ocMONEY)\r\n'
    b'    \'If vICMSAliq = "" Then Tb("vICMS") = CDbl(Format(0, "@")) Else Tb("vICMS") = CDbl(Format(vValorICMS, "@"))\r\n'
)

new = (
    b'    \' vBC e vICMS\r\n'
    b'    If vTipoCRT = 1 Then\r\n'
    b'        \' Simples Nacional: zera vBC/vICMS exceto CSOSN 101 e 201 (permite credito)\r\n'
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
    b'        \' Regime Normal: aplica reducao de BC se houver\r\n'
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

c = data.count(old)
if c == 1:
    data = data.replace(old, new)
    print('OK')
else:
    print(f'ERRO: {c} ocorrencias')
    # Debug: show what's actually there
    idx = data.find(b'    Tb("vBC") = CDbl(Format(txtSubTotal, "@"))')
    if idx >= 0:
        print('Encontrado em:', idx)
        print(repr(data[idx:idx+500]))

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
