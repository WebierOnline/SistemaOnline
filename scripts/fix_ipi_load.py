data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'    If vIPIALIQ = "" Then Tb("IPIvIPI") = CDbl(Format(0, "@")) Else Tb("IPIvIPI") = CDbl(Format(vValorIPI, "@"))\r\n'
    b'    \r\n'
    b'    \'Valores do item\r\n'
)
new = (
    b'    If vIPIALIQ = "" Then Tb("IPIvIPI") = CDbl(Format(0, "@")) Else Tb("IPIvIPI") = CDbl(Format(vValorIPI, "@"))\r\n'
    b'    Tb("IPICST") = Format(vIPICST, "@")\r\n'
    b'    If vIPICST = "99" Or vIPICST = "53" Or vIPICST = "52" Or vIPICST = "50" Then\r\n'
    b'        Tb("IPIcEnq") = "999"\r\n'
    b'    Else\r\n'
    b'        Tb("IPIcEnq") = ""\r\n'
    b'    End If\r\n'
    b'    \r\n'
    b'    \'Valores do item\r\n'
)

c = data.count(old)
if c == 1:
    data = data.replace(old, new)
    print('OK')
else:
    print(f'ERRO: {c} ocorrencias')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
