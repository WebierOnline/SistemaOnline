data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'Sub Exibir_Itens()\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'\r\n'
    b'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, pICMS, vICMS, ValorUnitarioComercializacao, QuantidadeComercial, valordesconto, ValorTotalBruto, IPIpIPI, IPIvIPI, ValorFrete, ValorSeguro, ValorOutros  FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'FormatarGridItensNota Tb\r\n'
    b'End Sub'
)
new = (
    b'Sub Exibir_Itens()\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'\r\n'
    b'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, " & _\r\n'
    b'       "ValorUnitarioComercializacao, QuantidadeComercial, ValorTotalBruto, " & _\r\n'
    b'       "ValorFrete, ValorSeguro, ValorOutros, ValorDesconto, " & _\r\n'
    b'       "vBC, pICMS, vICMS, pRedBC, " & _\r\n'
    b'       "vBCST, pICMSST, vICMSST, pMVAST, " & _\r\n'
    b'       "IPIpIPI, IPIvIPI, IPIcEnq " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'FormatarGridItensNota Tb\r\n'
    b'AplicarVisibilidadeGridItens\r\n'
    b'End Sub'
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
