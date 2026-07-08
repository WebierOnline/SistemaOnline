data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# Fix 1: main proportional UPDATE — add TipoDesconto and Desconto
old1 = (
    b'sSQL = "UPDATE NotaFiscalItens SET " & _\r\n'
    b'       "ValorDesconto = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2) " & _\r\n'
    b'       "WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'SQLExecuta sSQL'
)
new1 = (
    b'sSQL = "UPDATE NotaFiscalItens SET " & _\r\n'
    b'       "TipoDesconto = 1, " & _\r\n'
    b'       "Desconto     = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2), " & _\r\n'
    b'       "ValorDesconto = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2) " & _\r\n'
    b'       "WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'SQLExecuta sSQL'
)

# Fix 2: adjustment UPDATE TOP(1) — add Desconto alongside ValorDesconto
old2 = (
    b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & " " & _\r\n'
    b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _\r\n'
    b'           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _\r\n'
    b'           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"\r\n'
    b'    SQLExecuta sSQL'
)
new2 = (
    b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET " & _\r\n'
    b'           "ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & ", " & _\r\n'
    b'           "Desconto      = ValorDesconto + " & FSQL(vResto, 2) & " " & _\r\n'
    b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _\r\n'
    b'           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _\r\n'
    b'           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"\r\n'
    b'    SQLExecuta sSQL'
)

for i, (old, new) in enumerate([(old1, new1), (old2, new2)], 1):
    count = data.count(old)
    if count == 1:
        data = data.replace(old, new)
        print(f'{i} OK')
    else:
        print(f'{i} ERRO: encontrado {count} vezes')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
