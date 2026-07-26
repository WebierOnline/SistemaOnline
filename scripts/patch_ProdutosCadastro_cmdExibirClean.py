path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'
COD = 'C\xd3D. BARRA'

old = (
    '      "WHERE " & varProdutoHabilitado & " " & varTipoMostrar & " " & vUltimoValorVenda & "  and "' + R +
    '    End If' + R + R +
    '   If cboCriterios.Text = "' + COD + '" Then'
)

new = (
    '      "WHERE " & varProdutoHabilitado & " " & varTipoMostrar & " " & vUltimoValorVenda & "  and "' + R +
    '    End If' + R + R +
    '   If cboCriterios.Text = "' + COD + '" Then' + R +
    '       cboConsProduto.Text = Replace(cboConsProduto.Text, " ", "")' + R +
    '   ElseIf cboCriterios.Text = "NCM" Then' + R +
    '       Dim sNCM As String, iNCM As Integer' + R +
    '       sNCM = ""' + R +
    '       For iNCM = 1 To Len(cboConsProduto.Text)' + R +
    '           If Mid(cboConsProduto.Text, iNCM, 1) >= "0" And Mid(cboConsProduto.Text, iNCM, 1) <= "9" Then' + R +
    '               sNCM = sNCM & Mid(cboConsProduto.Text, iNCM, 1)' + R +
    '           End If' + R +
    '       Next iNCM' + R +
    '       cboConsProduto.Text = sNCM' + R +
    '   End If' + R + R +
    '   If cboCriterios.Text = "' + COD + '" Then'
)

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
