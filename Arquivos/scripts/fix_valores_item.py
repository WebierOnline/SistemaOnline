data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'    Tb("ValorFrete") = CDbl(Format(txtFrete, "@"))\r\n'
    b'    Tb("ValorSeguro") = CDbl(Format(txtSeguro, "@"))\r\n'
    b'    Tb("ValorOutros") = CDbl(Format(txtOutrosItem, "@"))\r\n'
)
new = (
    b'    Tb("ValorFrete") = CDbl(IIf(txtFrete.Text = "", 0, Format(txtFrete, "@")))\r\n'
    b'    Tb("ValorSeguro") = CDbl(IIf(txtSeguro.Text = "", 0, Format(txtSeguro, "@")))\r\n'
    b'    Tb("ValorOutros") = CDbl(IIf(txtOutrosItem.Text = "", 0, Format(txtOutrosItem, "@")))\r\n'
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
