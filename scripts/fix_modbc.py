data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# Adicionar modBC = 3 na secao 'ICMS, logo apos Tb("CST")
old = (
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
)
new = (
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    Tb("modBC") = Format(3, "@")\r\n'
    b'    If vICMSAliq = "" Then Tb("pICMS") = CDbl(Format(0, "@")) Else Tb("pICMS") = CDbl(Format(vICMSAliq, "@"))\r\n'
)

c = data.count(old)
print(f'modBC: {c}')
if c == 1: data = data.replace(old, new)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
