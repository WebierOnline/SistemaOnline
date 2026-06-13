data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1: reverter Load_Data e LerDadosInserir — remover flag bNFAgrupada ──
old1 = ('    If bNFAgrupada Then\r\n'
        '        TbNotas("Cod_Pedido") = 0\r\n'
        '    Else\r\n'
        '        TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")\r\n'
        '    End If\r\n')
new1 = '    TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")\r\n'

# ── Patch 2: no cmdConverterNFe_Click, após GravarPedido fazer UPDATE com 0 ───
old2 = ('                bNFAgrupada = True\r\n'
        '                vTipoEdicaoNFe = "Edicao"\r\n'
        '                GravarPedido\r\n'
        '                bNFAgrupada = False\r\n'
        '                bPrimeiro = False\r\n')
new2 = ('                vTipoEdicaoNFe = "Edicao"\r\n'
        '                GravarPedido\r\n'
        '                SQLExecuta "UPDATE NotaFiscal SET Cod_Pedido = 0 WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        '                bPrimeiro = False\r\n')

c1 = text.count(old1)
c2 = text.count(old2)
print(f'Patch1: {c1}  Patch2: {c2}')

if c1 == 2 and c2 == 1:
    text = text.replace(old1, new1)
    text = text.replace(old2, new2)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK')
else:
    print('ABORTADO')
