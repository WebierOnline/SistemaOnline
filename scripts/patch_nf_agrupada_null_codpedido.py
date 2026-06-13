data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

old = (
    '                vTipoEdicaoNFe = "Edicao"\r\n'
    '                GravarPedido\r\n'
    '                bPrimeiro = False\r\n'
    '                sPedidosLista = Format(CLng(txtCodPedido.Text), "0000")\r\n'
)
new = (
    '                vTipoEdicaoNFe = "Edicao"\r\n'
    '                GravarPedido\r\n'
    '                SQLExecuta "UPDATE NotaFiscal SET Cod_Pedido = NULL WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    '                bPrimeiro = False\r\n'
    '                sPedidosLista = Format(CLng(txtCodPedido.Text), "0000")\r\n'
)

c = text.count(old)
print('count:', c)
if c == 1:
    text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK')
else:
    print('ERRO')
