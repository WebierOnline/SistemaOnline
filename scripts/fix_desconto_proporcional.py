data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'Private Sub DistribuirDesconto()\r\n'
    b'Dim varQuantItens   As Integer\r\n'
    b'Dim vTotalDesc      As Currency\r\n'
    b'Dim vDescIndividual As Currency\r\n'
    b'Dim vDescAjuste     As Currency\r\n'
    b'\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'If GridNotasItens.rows <= 1 Then Exit Sub\r\n'
    b'\r\n'
    b'sSQL = "SELECT codigonota FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'If Not Tb.EOF Then\r\n'
    b'    varQuantItens = Tb.RecordCount\r\n'
    b'Else\r\n'
    b'    varQuantItens = 0\r\n'
    b'End If\r\n'
    b'\r\n'
    b'If txtValorDesconto.Text <> "0" And txtValorDesconto.Text <> "" Then\r\n'
    b'    vTotalDesc = txtValorDesconto.Text\r\n'
    b'Else\r\n'
    b'    vTotalDesc = 0\r\n'
    b'End If\r\n'
    b'\r\n'
    b'If vTotalDesc = 0 Or varQuantItens = 0 Then\r\n'
    b'    Exit Sub\r\n'
    b'Else\r\n'
    b'    vDescIndividual = CCur(Format(vTotalDesc / varQuantItens, "0.00"))\r\n'
    b'    vDescAjuste     = vTotalDesc - (vDescIndividual * (varQuantItens - 1))\r\n'
    b'    \r\n'
    b'    sSQL = "UPDATE NotaFiscalItens SET ValorDesconto = " & FSQL(vDescIndividual, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    SQLExecuta sSQL\r\n'
    b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorDesconto = " & FSQL(vDescAjuste, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    SQLExecuta sSQL\r\n'
    b'End If\r\n'
    b'End Sub'
)

new = (
    b'Private Sub DistribuirDesconto()\r\n'
    b'Dim vTotalDesc      As Currency\r\n'
    b'Dim vTotalSubtotal  As Currency\r\n'
    b'Dim vResto          As Currency\r\n'
    b'Dim rTot            As ADODB.Recordset\r\n'
    b'Dim rResto          As ADODB.Recordset\r\n'
    b'\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'If GridNotasItens.rows <= 1 Then Exit Sub\r\n'
    b'\r\n'
    b'If txtValorDesconto.Text <> "0" And txtValorDesconto.Text <> "" Then\r\n'
    b'    vTotalDesc = txtValorDesconto.Text\r\n'
    b'Else\r\n'
    b'    vTotalDesc = 0\r\n'
    b'End If\r\n'
    b'\r\n'
    b'If vTotalDesc = 0 Then Exit Sub\r\n'
    b'\r\n'
    b"' Busca subtotal total dos itens\r\n"
    b'sSQL = "SELECT ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS TotalSubtotal " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'Set rTot = dbData.OpenRecordset(sSQL)\r\n'
    b'If rTot.EOF Then Exit Sub\r\n'
    b'vTotalSubtotal = CCur(rTot("TotalSubtotal"))\r\n'
    b'If vTotalSubtotal = 0 Then Exit Sub\r\n'
    b'\r\n'
    b"' Valida: desconto nao pode exceder o subtotal total dos produtos\r\n"
    b'If vTotalDesc > vTotalSubtotal Then\r\n'
    b'    ShowMsg "O desconto total (" & FormatNumber(vTotalDesc, 2) & ") nao pode ser maior que o subtotal dos produtos (" & FormatNumber(vTotalSubtotal, 2) & ").", vbExclamation\r\n'
    b'    txtValorDesconto.Text = FormatNumber(vTotalSubtotal, 2)\r\n'
    b'    Exit Sub\r\n'
    b'End If\r\n'
    b'\r\n'
    b"' Distribui proporcionalmente ao subtotal de cada item\r\n"
    b'sSQL = "UPDATE NotaFiscalItens SET " & _\r\n'
    b'       "ValorDesconto = ROUND(" & FSQL(vTotalDesc, 2) & " * (ValorUnitarioComercializacao * QuantidadeComercial) / " & FSQL(vTotalSubtotal, 2) & ", 2) " & _\r\n'
    b'       "WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'SQLExecuta sSQL\r\n'
    b'\r\n'
    b"' Ajusta o resto do arredondamento no item com maior subtotal (tem mais margem)\r\n"
    b'sSQL = "SELECT " & FSQL(vTotalDesc, 2) & " - ISNULL(SUM(ValorDesconto), 0) AS Resto " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'Set rResto = dbData.OpenRecordset(sSQL)\r\n'
    b'If Not rResto.EOF Then vResto = CCur(rResto("Resto"))\r\n'
    b'\r\n'
    b'If vResto <> 0 Then\r\n'
    b'    sSQL = "UPDATE TOP(1) NotaFiscalItens SET ValorDesconto = ValorDesconto + " & FSQL(vResto, 2) & " " & _\r\n'
    b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " " & _\r\n'
    b'           "AND (ValorUnitarioComercializacao * QuantidadeComercial) = " & _\r\n'
    b'           "(SELECT MAX(ValorUnitarioComercializacao * QuantidadeComercial) FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & ")"\r\n'
    b'    SQLExecuta sSQL\r\n'
    b'End If\r\n'
    b'End Sub'
)

count = data.count(old)
if count == 1:
    data = data.replace(old, new)
    print('1 OK')
else:
    print(f'ERRO: encontrado {count} vezes')

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
