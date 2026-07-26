path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AdicionarQuant.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

CRLF = '\r\n'

# Inserir cmdExato_Click antes de cmdSair_Click
content = sub('add cmdExato_Click',
    'Private Sub cmdSair_Click()',

    'Private Sub cmdExato_Click()' + CRLF +
    'Dim sSQL As String' + CRLF +
    'Dim r As ADODB.Recordset' + CRLF +
    'Dim AutoNumeracao As Long' + CRLF +
    'Dim dExato As Double' + CRLF +
    'Dim dAtual As Double' + CRLF +
    'Dim dDiff As Double' + CRLF +
    'Dim sTipo As String' + CRLF +
    CRLF +
    'If Len(Trim(txtQuantNova.Text)) = 0 Or Not IsNumeric(txtQuantNova.Text) Then' + CRLF +
    '   MsgBox "Digite uma quantidade v\xe1lida.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dExato = Val(Replace(Replace(txtQuantNova.Text, ".", ""), ",", "."))' + CRLF +
    'If dExato < 0 Then' + CRLF +
    '   MsgBox "A quantidade n\xe3o pode ser negativa.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dAtual = Val(Replace(Replace(txtQuant.Text, ".", ""), ",", "."))' + CRLF +
    'dDiff = dExato - dAtual' + CRLF +
    CRLF +
    'If dDiff = 0 Then' + CRLF +
    '   MsgBox "O estoque j\xe1 est\xe1 em " & txtQuantNova.Text & ". Nenhuma altera\xe7\xe3o necess\xe1ria.", vbInformation, "Exato"' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    CRLF +
    'If dDiff > 0 Then' + CRLF +
    '   sTipo = "ADI\xc7\xc3O"' + CRLF +
    'Else' + CRLF +
    '   sTipo = "REMO\xc7\xc3O"' + CRLF +
    'End If' + CRLF +
    CRLF +
    'If ShowMsg("Deseja ajustar o estoque de " & txtDescricao.Text & " para " & txtQuantNova.Text & " unidade(s)?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub' + CRLF +
    CRLF +
    "'AUTONUMERA\xc7\xc3O" + CRLF +
    'sSQL = "SELECT ISNULL(MAX(codigo), 0) AS cod_itens FROM Produtos_Quant;"' + CRLF +
    'Set r = dbData.OpenRecordset(sSQL)' + CRLF +
    'If Not r.BOF Then AutoNumeracao = r("cod_itens") + 1' + CRLF +
    'If r.State <> 0 Then r.Close' + CRLF +
    'Set r = Nothing' + CRLF +
    CRLF +
    'sSQL = "INSERT INTO Produtos_Quant (Codigo, COD_PRODUTO, Data, COD_ENTRADA, FORMA, QUANT, TIPO, HORA, COD_USUARIO, ESTOQUE) VALUES (" & _' + CRLF +
    '   AutoNumeracao & ", " & txtCodProduto.Text & ", CONVERT(DATETIME, \'' + "'" + '" & Format(Date, ocDATA) & "\'' + "'" + ', 103), 0, ' + "'" + 'AJUSTE' + "'" + ', " & Replace(CStr(Abs(dDiff)), ",", ".") & ", ' + "'" + '" & sTipo & "' + "'" + ', ' + "'" + '" & Format(Now, ocHRMN) & "' + "'" + ', " & txtCodUsuario.Text & ", " & Replace(CStr(dExato), ",", ".") & ");"' + CRLF +
    'dbData.Execute sSQL' + CRLF +
    CRLF +
    "'Ajusta o estoque para o valor exato" + CRLF +
    'dbData.Execute "UPDATE produtos SET quant_estoque = " & Replace(CStr(dExato), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"' + CRLF +
    CRLF +
    'cmdSair_Click' + CRLF +
    'End Sub' + CRLF +
    CRLF +
    'Private Sub cmdSair_Click()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
