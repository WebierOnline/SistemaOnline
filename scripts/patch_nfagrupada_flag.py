data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1: declarar bNFAgrupada como variável de módulo ─────────────────────
old1 = 'Attribute VB_Exposed = False\r\nPrivate moCombo As cComboHelper\r\n'
new1 = 'Attribute VB_Exposed = False\r\nPrivate moCombo As cComboHelper\r\nPrivate bNFAgrupada As Boolean\r\n'

# ── Patch 2: em Load_Data, usar Null quando bNFAgrupada = True ────────────────
old2 = '    TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")\r\n'
new2 = ('    If bNFAgrupada Then\r\n'
        '        TbNotas("Cod_Pedido") = Null\r\n'
        '    Else\r\n'
        '        TbNotas("Cod_Pedido") = Format(txtCodPedido.Text, "@")\r\n'
        '    End If\r\n')

# ── Patch 3: no modo AGRUPADA, ligar/desligar a flag em torno de GravarPedido ─
old3 = ('                vTipoEdicaoNFe = "Edicao"\r\n'
        '                GravarPedido\r\n'
        '                SQLExecuta "UPDATE NotaFiscal SET Cod_Pedido = NULL WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        '                bPrimeiro = False\r\n')
new3 = ('                bNFAgrupada = True\r\n'
        '                vTipoEdicaoNFe = "Edicao"\r\n'
        '                GravarPedido\r\n'
        '                bNFAgrupada = False\r\n'
        '                bPrimeiro = False\r\n')

c1 = text.count(old1)
c2 = text.count(old2)
c3 = text.count(old3)
print(f'Patch1: {c1}  Patch2: {c2}  Patch3: {c3}')

if c1 == 1 and c2 == 2 and c3 == 1:
    text = text.replace(old1, new1)
    text = text.replace(old2, new2)
    text = text.replace(old3, new3)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK')
else:
    print('ABORTADO')
