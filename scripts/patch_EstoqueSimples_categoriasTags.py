path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_Estoque_Simples.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# 1: cmdEdicaoColetiva_Click — chamar InserirTagSeNova apos UPDATE no loop principal
content = sub('cmdEdicaoColetiva InserirTagSeNova',
    '   If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProximaLinha\r\n'
    '   If iColE >= 0 Then Grid.TextMatrix(i, iColE) = sVal\r\n'
    '   dbData.Execute "UPDATE produtos SET " & sSet & " WHERE codigo = " & Grid.TextMatrix(i, 2) & ";"\r\n'
    'ProximaLinha:\r\n'
    'Next i',

    '   If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProximaLinha\r\n'
    '   If iColE >= 0 Then Grid.TextMatrix(i, iColE) = sVal\r\n'
    '   dbData.Execute "UPDATE produtos SET " & sSet & " WHERE codigo = " & Grid.TextMatrix(i, 2) & ";"\r\n'
    '   If optETags.Value Then InserirTagSeNova Grid.TextMatrix(i, 7), sVal\r\n'
    'ProximaLinha:\r\n'
    'Next i',
    content)

# 2: cmdAtualizar_Click — a) validacao antes do loop  b) InserirTagSeNova dentro do loop
content = sub('cmdAtualizar validacao + InserirTagSeNova',
    'picAguarde.Visible = True\r\n'
    'DoEvents\r\n'
    '    \'txtDescricao.Text = TirarEspaco(txtDescricao.Text)\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   \'Atualiza a tabela de produtos\r\n'
    '   sSQL = "UPDATE produtos SET " & _\r\n'
    '      "cod_barra = \'' + "'" + ' & Grid.TextMatrix(i, 3) & "\', " & _\r\n'
    '      "descricao = \'' + "'" + ' & TirarEspaco(Grid.TextMatrix(i, 4)) & "\', " & _\r\n'
    '      "UNID_MEDIDA = \'' + "'" + ' & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "categoria = \'' + "'" + ' & Grid.TextMatrix(i, 7) & "\', " & _\r\n'
    '      "TAGS = \'' + "'" + ' & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "fabricante = \'' + "'" + ' & Grid.TextMatrix(i, 5) & "\', " & _\r\n'
    '      "PRATELEIRA = \'' + "'" + ' & Grid.TextMatrix(i, 9) & "\', " & _\r\n'
    '      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 10)), ",", ".") & " " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',

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
    '   \'Atualiza a tabela de produtos\r\n'
    '   sSQL = "UPDATE produtos SET " & _\r\n'
    '      "cod_barra = \'' + "'" + ' & Grid.TextMatrix(i, 3) & "\', " & _\r\n'
    '      "descricao = \'' + "'" + ' & TirarEspaco(Grid.TextMatrix(i, 4)) & "\', " & _\r\n'
    '      "UNID_MEDIDA = \'' + "'" + ' & Grid.TextMatrix(i, 6) & "\', " & _\r\n'
    '      "categoria = \'' + "'" + ' & Grid.TextMatrix(i, 7) & "\', " & _\r\n'
    '      "TAGS = \'' + "'" + ' & Grid.TextMatrix(i, 8) & "\', " & _\r\n'
    '      "fabricante = \'' + "'" + ' & Grid.TextMatrix(i, 5) & "\', " & _\r\n'
    '      "PRATELEIRA = \'' + "'" + ' & Grid.TextMatrix(i, 9) & "\', " & _\r\n'
    '      "ESTOQUE_FISCAL = " & Replace(CDbl(Grid.TextMatrix(i, 10)), ",", ".") & " " & _\r\n'
    '      "WHERE (codigo = " & Grid.TextMatrix(i, 2) & ");"',
    content)

# 3: cmdAtualizar_Click — apos dbData.Execute sSQL, inserir tag se nova
content = sub('cmdAtualizar apos Execute InserirTagSeNova',
    '      \'Debug.Print sSQL\r\n'
    '   dbData.Execute sSQL\r\n'
    'Next\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'cmdLocalizar_Click\r\n'
    'End Sub',

    '      \'Debug.Print sSQL\r\n'
    '   dbData.Execute sSQL\r\n'
    '   If Len(Trim(Grid.TextMatrix(i, 8))) > 0 Then\r\n'
    '      InserirTagSeNova Grid.TextMatrix(i, 7), Trim(Grid.TextMatrix(i, 8))\r\n'
    '   End If\r\n'
    'Next\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'cmdLocalizar_Click\r\n'
    'End Sub',
    content)

# 4: Inserir sub InserirTagSeNova antes de Form_Unload
content = sub('add InserirTagSeNova sub',
    'Private Sub Form_Unload(Cancel As Integer)\r\n'
    '   Set moCombo = Nothing\r\n'
    'End Sub',

    'Private Sub InserirTagSeNova(sCat As String, sTag As String)\r\n'
    'Dim sTagU As String\r\n'
    'Dim nID As Long\r\n'
    'Dim nExiste As Long\r\n'
    'sTagU = UCase(Trim(sTag))\r\n'
    'If Len(sTagU) = 0 Or Len(Trim(sCat)) = 0 Then Exit Sub\r\n'
    'nID = Val(SQLExecutaRetorno("SELECT ID_Categoria FROM Categorias WHERE Categoria = \'' + "'" + ' & Replace(sCat, "\'", "\'\'") & "\' AND Tipo_Empresa = " & tipoEmpresa, "ID_Categoria"))\r\n'
    'If nID = 0 Then Exit Sub\r\n'
    'nExiste = Val(SQLExecutaRetorno("SELECT COUNT(*) AS n FROM Categorias_Tags WHERE ID_Categoria = " & nID & " AND Tags = \'' + "'" + ' & Replace(sTagU, "\'", "\'\'") & "\'", "n"))\r\n'
    'If nExiste = 0 Then\r\n'
    '   dbData.Execute "INSERT INTO Categorias_Tags (ID_Categoria, Tags) VALUES (" & nID & ", \'' + "'" + ' & Replace(sTagU, "\'", "\'\'") & "\');"\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub Form_Unload(Cancel As Integer)\r\n'
    '   Set moCombo = Nothing\r\n'
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
