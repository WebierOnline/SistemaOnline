data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# Each Distribuir* ends with:
#   SQLExecuta sSQL   (the TOP(1) adjustment)
# End If
# End Sub
#
# We add Exibir_Itens before End If / End Sub

fixes = [
    # DistribuirFrete
    (
        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET valorfrete = " & FSQL(varValorAjusteFrete, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'End If\r\n'
        b'End Sub',

        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET valorfrete = " & FSQL(varValorAjusteFrete, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'    Exibir_Itens\r\n'
        b'End If\r\n'
        b'End Sub'
    ),
    # DistribuirSeguro
    (
        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorSeguro = " & FSQL(vValorSeguroAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'End If\r\n'
        b'End Sub',

        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorSeguro = " & FSQL(vValorSeguroAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'    Exibir_Itens\r\n'
        b'End If\r\n'
        b'End Sub'
    ),
    # DistribuirOutros
    (
        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorOutros = " & FSQL(vValorOutrosAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'End If\r\n'
        b'End Sub',

        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorOutros = " & FSQL(vValorOutrosAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'    Exibir_Itens\r\n'
        b'End If\r\n'
        b'End Sub'
    ),
    # DistribuirDesconto - ends differently (no TOP(1) fixed block, uses vResto)
    (
        b'If vResto <> 0 Then\r\n'
        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET " & _\r\n'
        b'           "ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & ", " & _\r\n'
        b'           "Desconto      = ValorDesconto + " & FSQL(vResto, 2) & " " & _\r\n'
        b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _\r\n'
        b'           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _\r\n'
        b'           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'End If\r\n'
        b'End Sub',

        b'If vResto <> 0 Then\r\n'
        b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET " & _\r\n'
        b'           "ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & ", " & _\r\n'
        b'           "Desconto      = ValorDesconto + " & FSQL(vResto, 2) & " " & _\r\n'
        b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _\r\n'
        b'           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _\r\n'
        b'           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"\r\n'
        b'    SQLExecuta sSQL\r\n'
        b'End If\r\n'
        b'Exibir_Itens\r\n'
        b'End Sub'
    ),
]

for i, (old, new) in enumerate(fixes, 1):
    count = data.count(old)
    if count == 1:
        data = data.replace(old, new)
        print(f'{i} OK')
    else:
        print(f'{i} ERRO: encontrado {count} vezes')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
