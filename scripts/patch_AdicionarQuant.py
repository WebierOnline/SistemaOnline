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
q    = chr(39)   # aspas simples — evita conflito com strings Python

# ── 1: cmdAdicionar — adicionar declaracoes + validacao ──────────────────────
content = sub('cmdAdicionar validacao',
    'Private Sub cmdAdicionar_Click()' + CRLF +
    'Dim sSQL As String' + CRLF +
    'Dim r As ADODB.Recordset' + CRLF +
    'Dim AutoNumeracao As Long' + CRLF +
    CRLF +
    "'AUTONUMERA\xc7\xc3O" + CRLF,

    'Private Sub cmdAdicionar_Click()' + CRLF +
    'Dim sSQL As String' + CRLF +
    'Dim r As ADODB.Recordset' + CRLF +
    'Dim AutoNumeracao As Long' + CRLF +
    'Dim dNova As Double' + CRLF +
    'Dim dAtual As Double' + CRLF +
    CRLF +
    'If Len(Trim(txtQuantNova.Text)) = 0 Or Not IsNumeric(txtQuantNova.Text) Then' + CRLF +
    '   MsgBox "Digite uma quantidade v\xe1lida.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dNova = Val(Replace(Replace(txtQuantNova.Text, ".", ""), ",", "."))' + CRLF +
    'If dNova <= 0 Then' + CRLF +
    '   MsgBox "A quantidade deve ser maior que zero.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dAtual = Val(Replace(Replace(txtQuant.Text, ".", ""), ",", "."))' + CRLF +
    CRLF +
    "'AUTONUMERA\xc7\xc3O" + CRLF,
    content)

# ── 2: cmdAdicionar — fix ESTOQUE no INSERT (saldo apos adicao) ──────────────
content = sub('cmdAdicionar ESTOQUE fix',
    'txtCodUsuario.Text & ", " & Replace(CDbl(txtQuant.Text), ",", ".") & ");"' + CRLF +
    'dbData.Execute sSQL' + CRLF +
    CRLF +
    "'Atualiza o estoque do produto" + CRLF +
    'dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(txtQuantNova.Text, ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"',

    'txtCodUsuario.Text & ", " & Replace(CStr(dAtual + dNova), ",", ".") & ");"' + CRLF +
    'dbData.Execute sSQL' + CRLF +
    CRLF +
    "'Atualiza o estoque do produto" + CRLF +
    'dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque + " & Replace(CStr(dNova), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"',
    content)

# ── 3: cmdRemover — adicionar declaracoes + validacao + cheque negativo ───────
content = sub('cmdRemover validacao',
    'Private Sub cmdRemover_Click()' + CRLF +
    'Dim sSQL As String' + CRLF +
    'Dim r As ADODB.Recordset' + CRLF +
    'Dim AutoNumeracao As Long' + CRLF +
    CRLF +
    "'AUTONUMERA\xc7\xc3O" + CRLF,

    'Private Sub cmdRemover_Click()' + CRLF +
    'Dim sSQL As String' + CRLF +
    'Dim r As ADODB.Recordset' + CRLF +
    'Dim AutoNumeracao As Long' + CRLF +
    'Dim dNova As Double' + CRLF +
    'Dim dAtual As Double' + CRLF +
    CRLF +
    'If Len(Trim(txtQuantNova.Text)) = 0 Or Not IsNumeric(txtQuantNova.Text) Then' + CRLF +
    '   MsgBox "Digite uma quantidade v\xe1lida.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dNova = Val(Replace(Replace(txtQuantNova.Text, ".", ""), ",", "."))' + CRLF +
    'If dNova <= 0 Then' + CRLF +
    '   MsgBox "A quantidade deve ser maior que zero.", vbExclamation, "Quantidade inv\xe1lida"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    'dAtual = Val(Replace(Replace(txtQuant.Text, ".", ""), ",", "."))' + CRLF +
    'If dNova > dAtual Then' + CRLF +
    '   MsgBox "Quantidade a remover (" & txtQuantNova.Text & ") \xe9 maior que o estoque atual (" & txtQuant.Text & ").", vbExclamation, "Estoque insuficiente"' + CRLF +
    '   txtQuantNova.SetFocus' + CRLF +
    '   Exit Sub' + CRLF +
    'End If' + CRLF +
    CRLF +
    "'AUTONUMERA\xe7\xc3O" + CRLF,
    content)

# ── 4: cmdRemover — fix ESTOQUE no INSERT (saldo apos remocao) ───────────────
content = sub('cmdRemover ESTOQUE fix',
    'txtCodUsuario.Text & ", " & Replace(CDbl(txtQuant.Text), ",", ".") & ");"' + CRLF +
    'dbData.Execute sSQL' + CRLF +
    CRLF +
    "'Atualiza o estoque do produto" + CRLF +
    'dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(txtQuantNova.Text, ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"',

    'txtCodUsuario.Text & ", " & Replace(CStr(dAtual - dNova), ",", ".") & ");"' + CRLF +
    'dbData.Execute sSQL' + CRLF +
    CRLF +
    "'Atualiza o estoque do produto" + CRLF +
    'dbData.Execute "UPDATE produtos SET quant_estoque = quant_estoque - " & Replace(CStr(dNova), ",", ".") & " WHERE (codigo = " & txtCodProduto.Text & ");"',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
