path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

# --- Change 1: cboConsLinha_GotFocus dinamico ---
old1 = (
    'Private Sub cboConsLinha_GotFocus()\r\n'
    '   Dim sSQL As String\r\n'
    '   Dim r As ADODB.Recordset\r\n'
    '   \r\n'
    '   cboConsLinha.Clear\r\n'
    '   \r\n'
    '   sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"\r\n'
    '   Set r = dbData.OpenRecordset(sSQL)\r\n'
    '   \r\n'
    '   Do While Not r.EOF\r\n'
    '      cboConsLinha.AddItem ValidateNull(r("categoria"))\r\n'
    '      r.MoveNext\r\n'
    '   Loop\r\n'
    '   \r\n'
    '   If r.State <> 0 Then r.Close\r\n'
    '   Set r = Nothing\r\n'
    '   \r\n'
    '   moCombo.AttachTo cboConsLinha\r\n'
    'End Sub'
)

new1 = (
    'Private Sub cboConsLinha_GotFocus()\r\n'
    '   Dim sSQL As String\r\n'
    '   Dim r As ADODB.Recordset\r\n'
    '   Dim sField As String\r\n'
    '   \r\n'
    '   cboConsLinha.Clear\r\n'
    '   \r\n'
    '   If optCategoria.Value Then\r\n'
    '      sField = "categoria"\r\n'
    '      sSQL = "SELECT DISTINCT categoria FROM produtos ORDER BY categoria;"\r\n'
    '   ElseIf optTags.Value Then\r\n'
    '      sField = "TAGS"\r\n'
    "      sSQL = \"SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' ORDER BY TAGS;\"\r\n"
    '   ElseIf optNCM.Value Then\r\n'
    '      sField = "NCM"\r\n'
    "      sSQL = \"SELECT DISTINCT NCM FROM produtos WHERE NCM IS NOT NULL AND NCM <> '' ORDER BY NCM;\"\r\n"
    '   ElseIf optClassTribCBS.Value Then\r\n'
    '      sField = "val"\r\n'
    "      sSQL = \"SELECT cClassTrib + ' - ' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;\"\r\n"
    '   ElseIf optClassTribIS.Value Then\r\n'
    '      sField = "val"\r\n'
    "      sSQL = \"SELECT cClassTrib_IS + ' - ' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;\"\r\n"
    '   Else\r\n'
    '      Exit Sub\r\n'
    '   End If\r\n'
    '   \r\n'
    '   Set r = dbData.OpenRecordset(sSQL)\r\n'
    '   \r\n'
    '   Do While Not r.EOF\r\n'
    '      cboConsLinha.AddItem ValidateNull(r(sField))\r\n'
    '      r.MoveNext\r\n'
    '   Loop\r\n'
    '   \r\n'
    '   If r.State <> 0 Then r.Close\r\n'
    '   Set r = Nothing\r\n'
    '   \r\n'
    '   moCombo.AttachTo cboConsLinha\r\n'
    'End Sub'
)

c1 = content.count(old1)
if c1 != 1:
    errors.append(f'Change1: count={c1}')
else:
    content = content.replace(old1, new1, 1)
    print('Change 1 OK')

# --- Change 2: optCategoria_Click — reset caption ---
old2 = (
    'Private Sub optCategoria_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'cboConsLinha.Visible = True\r\n'
)
new2 = (
    'Private Sub optCategoria_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Categoria"\r\n'
    'cboConsLinha.Visible = True\r\n'
)
c2 = content.count(old2)
if c2 != 1:
    errors.append(f'Change2: count={c2}')
else:
    content = content.replace(old2, new2, 1)
    print('Change 2 OK')

# --- Change 3: WHERE clauses para os 4 novos filtros ---
old3 = (
    '& "produtos.cod_barra = \'" & txtCodBarra.Text & "\'", "")\r\n'
    '   \r\n'
    '   If var_Criterio <> "" Then var_Criterio = " WHERE " & var_Criterio'
)
new3 = (
    '& "produtos.cod_barra = \'" & txtCodBarra.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optTags.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.TAGS = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optNCM.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.NCM = \'" & cboConsLinha.Text & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optClassTribCBS.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib = \'" & Left(cboConsLinha.Text, 6) & "\'", "")\r\n'
    '   var_Criterio = var_Criterio & IIf(optClassTribIS.Value, IIf(var_Criterio <> "", " AND ", "") & "produtos.cClassTrib_IS = \'" & Left(cboConsLinha.Text, 6) & "\'", "")\r\n'
    '   \r\n'
    '   If var_Criterio <> "" Then var_Criterio = " WHERE " & var_Criterio'
)
c3 = content.count(old3)
if c3 != 1:
    errors.append(f'Change3: count={c3}')
else:
    content = content.replace(old3, new3, 1)
    print('Change 3 OK')

# --- Change 4: 4 novos handlers apos optTodosPreco_Click ---
new_handlers = (
    '\r\n'
    'Private Sub optTags_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Tags"\r\n'
    'cboConsLinha.Visible = True\r\n'
    'lblDesc.Visible = False\r\n'
    'cboDesc.Visible = False\r\n'
    'chkDescPorProduto.Visible = False\r\n'
    'chkDescPorIniciais.Visible = False\r\n'
    'lblCodBarra.Visible = False\r\n'
    'txtCodBarra.Visible = False\r\n'
    'cmdLocalizar.Visible = True\r\n'
    'Frame6.Enabled = False\r\n'
    'optTodosPreco.Value = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optNCM_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "NCM"\r\n'
    'cboConsLinha.Visible = True\r\n'
    'lblDesc.Visible = False\r\n'
    'cboDesc.Visible = False\r\n'
    'chkDescPorProduto.Visible = False\r\n'
    'chkDescPorIniciais.Visible = False\r\n'
    'lblCodBarra.Visible = False\r\n'
    'txtCodBarra.Visible = False\r\n'
    'cmdLocalizar.Visible = True\r\n'
    'Frame6.Enabled = False\r\n'
    'optTodosPreco.Value = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribCBS_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. CBS"\r\n'
    'cboConsLinha.Visible = True\r\n'
    'lblDesc.Visible = False\r\n'
    'cboDesc.Visible = False\r\n'
    'chkDescPorProduto.Visible = False\r\n'
    'chkDescPorIniciais.Visible = False\r\n'
    'lblCodBarra.Visible = False\r\n'
    'txtCodBarra.Visible = False\r\n'
    'cmdLocalizar.Visible = True\r\n'
    'Frame6.Enabled = False\r\n'
    'optTodosPreco.Value = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optClassTribIS_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. IS"\r\n'
    'cboConsLinha.Visible = True\r\n'
    'lblDesc.Visible = False\r\n'
    'cboDesc.Visible = False\r\n'
    'chkDescPorProduto.Visible = False\r\n'
    'chkDescPorIniciais.Visible = False\r\n'
    'lblCodBarra.Visible = False\r\n'
    'txtCodBarra.Visible = False\r\n'
    'cmdLocalizar.Visible = True\r\n'
    'Frame6.Enabled = False\r\n'
    'optTodosPreco.Value = True\r\n'
    'cboConsLinha.SetFocus\r\n'
    'End Sub\r\n'
)

old4 = (
    'Private Sub optTodosPreco_Click()\r\n'
    'cmdLocalizar_Click\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub txtCodBarra_Change()'
)
new4 = (
    'Private Sub optTodosPreco_Click()\r\n'
    'cmdLocalizar_Click\r\n'
    'End Sub\r\n'
    + new_handlers +
    '\r\n'
    'Private Sub txtCodBarra_Change()'
)
c4 = content.count(old4)
if c4 != 1:
    errors.append(f'Change4: count={c4}')
else:
    content = content.replace(old4, new4, 1)
    print('Change 4 OK')

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('File written successfully.')
