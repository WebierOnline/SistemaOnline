path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
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

R = '\r\n'
q = chr(39)   # aspas simples para o SQL

# helper: campo texto com escape de aspas simples
def qf(expr):
    return q + '" & Replace(' + expr + ', "' + q + '", "' + q + q + '") & "' + q

# helper: numero locale-safe para SQL (Val + CStr + Replace)
def nf(expr):
    return '" & Replace(CStr(Val(Replace(Replace(' + expr + ', ".", ""), ",", "."))), ",", ".") & "'

# ── Substituir todo o corpo do For no cmdAtualizar ───────────────────────────
old_for = (
    '   ' + q + 'Atualiza a tabela de produtos' + R +
    '   sSQL = "UPDATE produtos SET " & _' + R +
    '      "cod_barra = ' + q + '" & Grid.TextMatrix(i, 3) & "' + q + ', " & _' + R +
    '      "descricao = ' + q + '" & TirarEspaco(Grid.TextMatrix(i, 4)) & "' + q + ', " & _' + R +
    '      "UNID_MEDIDA = ' + q + '" & Grid.TextMatrix(i, 6) & "' + q + ', " & _' + R +
    '      "categoria = ' + q + '" & Grid.TextMatrix(i, 7) & "' + q + ', " & _' + R +
    '      "TAGS = ' + q + '" & Grid.TextMatrix(i, 8) & "' + q + ', " & _' + R +
    '      "fabricante = ' + q + '" & Grid.TextMatrix(i, 5) & "' + q + ', " & _' + R +
    '      "PRATELEIRA = ' + q + '" & Grid.TextMatrix(i, 9) & "' + q + ', " & _' + R +
    '      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 10)), ",", ".") & " " & _' + R +
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"' + R +
    '            ' + R +
    R +
    '      ' + q + '" & Replace(CDbl(Grid.TextMatrix(i, 15)), ",", ".") & "' + R +
    '      ' + q + '"ESTOQUE_FISCAL = ' + q + '" & Grid.TextMatrix(i, 9) & "' + q + ' " & _' + R +
    '    ' + q + 'sSQL = "UPDATE produtos SET " & _' + R +
    '      "cod_barra = ' + q + '" & Grid.TextMatrix(i, 3) & "' + q + ', " & _' + R +
    '      "descricao = ' + q + '" & Grid.TextMatrix(i, 4) & "' + q + ', " & _' + R +
    '      "UNID_MEDIDA = ' + q + '" & Grid.TextMatrix(i, 6) & "' + q + ', " & _' + R +
    '      "categoria = ' + q + '" & Grid.TextMatrix(i, 7) & "' + q + ', " & _' + R +
    '      "fabricante = ' + q + '" & Grid.TextMatrix(i, 5) & "' + q + ', " & _' + R +
    '      "PRATELEIRA = ' + q + '" & Grid.TextMatrix(i, 8) & "' + q + ', " & _' + R +
    '      "ESTOQUE_FISCAL = ' + q + '" & Grid.TextMatrix(i, 9) & "' + q + ' " & _' + R +
    '      "WHERE (codigo = 2057);"' + R +
    '      ' + q + 'Debug.Print sSQL' + R +
    '   dbData.Execute sSQL' + R +
    '   If Len(Trim(Grid.TextMatrix(i, 8))) > 0 Then' + R +
    '      InserirTagSeNova Grid.TextMatrix(i, 7), Trim(Grid.TextMatrix(i, 8))' + R +
    '   End If' + R +
    'Next'
)

new_for = (
    '   ' + q + 'Atualiza a tabela de produtos' + R +
    '   sSQL = "UPDATE produtos SET " & _' + R +
    '      "cod_barra = '    + qf('Trim(Grid.TextMatrix(i, 3))')         + ', " & _' + R +
    '      "descricao = '    + qf('TirarEspaco(Grid.TextMatrix(i, 4))') + ', " & _' + R +
    '      "UNID_MEDIDA = '  + qf('Grid.TextMatrix(i, 6)')               + ', " & _' + R +
    '      "categoria = '    + qf('Grid.TextMatrix(i, 7)')               + ', " & _' + R +
    '      "TAGS = '         + qf('Grid.TextMatrix(i, 8)')               + ', " & _' + R +
    '      "fabricante = '   + qf('Grid.TextMatrix(i, 5)')               + ', " & _' + R +
    '      "PRATELEIRA = '   + qf('Grid.TextMatrix(i, 9)')               + '"' + R +
    '   If optMostrarFiscal.Value Then' + R +
    '      sSQL = sSQL & ", ESTOQUE_FISCAL = ' + nf('Grid.TextMatrix(i, 10)') + '"' + R +
    '   Else' + R +
    '      sSQL = sSQL & ", quant_estoque = ' + nf('Grid.TextMatrix(i, 10)') + '"' + R +
    '   End If' + R +
    '   sSQL = sSQL & " WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"' + R +
    '   dbData.Execute sSQL' + R +
    '   If Len(Trim(Grid.TextMatrix(i, 8))) > 0 Then' + R +
    '      InserirTagSeNova Grid.TextMatrix(i, 7), Trim(Grid.TextMatrix(i, 8))' + R +
    '   End If' + R +
    'Next'
)

content = sub('cmdAtualizar rewrite loop', old_for, new_for, content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
