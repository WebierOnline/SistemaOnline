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

# cmdAtualizar_Click — validacao tag sem categoria antes do loop principal
content = sub('cmdAtualizar validacao tag',
    'picAguarde.Visible = True\r\n'
    'DoEvents\r\n'
    '    \'txtDescricao.Text = TirarEspaco(txtDescricao.Text)\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   \'Atualiza a tabela de produtos',

    'picAguarde.Visible = True\r\n'
    'DoEvents\r\n'
    '\r\n'
    'Dim iSemCatTag As Integer\r\n'
    'Dim jv As Integer\r\n'
    'For jv = 1 To Grid.rows - 1\r\n'
    '   If Len(Trim(Grid.TextMatrix(jv, 8))) > 0 And Len(Trim(Grid.TextMatrix(jv, 7))) = 0 Then\r\n'
    '      iSemCatTag = iSemCatTag + 1\r\n'
    '   End If\r\n'
    'Next jv\r\n'
    'If iSemCatTag > 0 Then\r\n'
    '   MsgBox iSemCatTag & " produto(s) com tag mas sem categoria. Defina a categoria antes de salvar.", vbExclamation, "Tag sem categoria"\r\n'
    '   picAguarde.Visible = False\r\n'
    '   Exit Sub\r\n'
    'End If\r\n'
    '\r\n'
    '    \'txtDescricao.Text = TirarEspaco(txtDescricao.Text)\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   \'Atualiza a tabela de produtos',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
