data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = b'    If vTipoCRT = 1 And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
new = b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'

c = data.count(old)
print(f'Ocorrencias: {c}')
if c == 1: data = data.replace(old, new)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
