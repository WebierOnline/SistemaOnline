data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'            dblAliqInter = CDbl(IIf(vAliqUFInter = 0, 0, vAliqUFInter))\r\n'
    b'\r\n'
    b"            ' MVA: ajustado se interestadual, original se interna\r\n"
)
new = (
    b'            If vUFEmpresa <> vUFDest And vUFDest <> "" Then\r\n'
    b'                Set rDifalLDI = dbData.OpenRecordset("SELECT AliquotaInterestadual FROM TribMatrizInterestadual WHERE UF_Origem = \'" & vUFEmpresa & "\' AND UF_Destino = \'" & vUFDest & "\'")\r\n'
    b'                If Not rDifalLDI.EOF Then\r\n'
    b'                    dblAliqInter = CDbl(rDifalLDI("AliquotaInterestadual"))\r\n'
    b'                Else\r\n'
    b'                    dblAliqInter = CDbl(IIf(vAliqUFInter = 0, 0, vAliqUFInter))\r\n'
    b'                End If\r\n'
    b'                rDifalLDI.Close\r\n'
    b'            Else\r\n'
    b'                dblAliqInter = 0\r\n'
    b'            End If\r\n'
    b'\r\n'
    b"            ' MVA: ajustado se interestadual, original se interna\r\n"
)

c = data.count(old)
print(f'count: {c}')
if c == 1:
    data = data.replace(old, new)
    data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
    print('Salvo. Tamanho:', len(data))
else:
    print('ERRO: trecho nao encontrado ou ambiguo')
