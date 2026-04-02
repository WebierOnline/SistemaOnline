data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = b'    If vTipoCRT = 1 Then\r\n'
new = b'    If vTipoCRT = 1 And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'

c = data.count(old)
if c == 1:
    data = data.replace(old, new)
    print('OK')
else:
    print(f'ERRO: {c} ocorrencias')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
