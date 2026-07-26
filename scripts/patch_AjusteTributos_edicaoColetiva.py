path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    c = c.replace(old, new, 1)
    print(f'{label} OK')
    return c

# ── 1: optTodos_Click — desabilitar frames antes do End Sub ────────────────
content = sub('optTodos disable frames',
    'optComPreco.Value = True\r\ncmdLocalizar_Click\r\nEnd Sub',
    'optComPreco.Value = True\r\ncmdLocalizar_Click\r\nfrmEdicao.Enabled = False\r\nfrmEdicaoFiltros.Enabled = False\r\nEnd Sub',
    content)

# ── 2: optCodBarra_Click — desabilitar frames antes do End Sub ─────────────
content = sub('optCodBarra disable frames',
    'optTodosPreco.Value = True\r\ntxtCodBarra.SetFocus\r\nEnd Sub',
    'optTodosPreco.Value = True\r\ntxtCodBarra.SetFocus\r\nfrmEdicao.Enabled = False\r\nfrmEdicaoFiltros.Enabled = False\r\nEnd Sub',
    content)

# ── 3: optDesc_Click — desabilitar frames antes do End Sub ────────────────
content = sub('optDesc disable frames',
    'optTodosPreco.Value = True\r\ncboDesc.SetFocus\r\nEnd Sub',
    'optTodosPreco.Value = True\r\ncboDesc.SetFocus\r\nfrmEdicao.Enabled = False\r\nfrmEdicaoFiltros.Enabled = False\r\nEnd Sub',
    content)

# ── 4: cmdLocalizar_Click — habilitar/desabilitar frames ao final ──────────
content = sub('cmdLocalizar enable/disable frames',
    'If optCodBarra.Value = True Then txtCodBarra_GotFocus\r\nEnd Sub',
    'If optCodBarra.Value = True Then txtCodBarra_GotFocus\r\n'
    'If optTodos.Value Or optCodBarra.Value Or optDesc.Value Then\r\n'
    '   frmEdicao.Enabled = False\r\n'
    '   frmEdicaoFiltros.Enabled = False\r\n'
    'Else\r\n'
    '   frmEdicao.Enabled = True\r\n'
    '   frmEdicaoFiltros.Enabled = True\r\n'
    'End If\r\n'
    'End Sub',
    content)

# ── 5: Inserir handlers optE* + cmdEdicaoColetiva antes do Form_Unload ─────
new_handlers = (
    'Private Sub optEICMS_Click()\r\n'
    '   lblEdicaoColetiva.Caption = "ICMS CST"\r\n'
    '   txtEdicaoColetiva.Visible = True\r\n'
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optENCM_Click()\r\n'
    '   lblEdicaoColetiva.Caption = "NCM"\r\n'
    '   txtEdicaoColetiva.Visible = True\r\n'
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECFOP_Click()\r\n'
    '   lblEdicaoColetiva.Caption = "CFOP"\r\n'
    '   txtEdicaoColetiva.Visible = True\r\n'
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECest_Click()\r\n'
    '   lblEdicaoColetiva.Caption = "CEST"\r\n'
    '   txtEdicaoColetiva.Visible = True\r\n'
    '   cboEdicaoColetiva.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECategoria_Click()\r\n'
    'Dim r As ADODB.Recordset\r\n'
    '   lblEdicaoColetiva.Caption = "Categoria"\r\n'
    '   txtEdicaoColetiva.Visible = False\r\n'
    '   cboEdicaoColetiva.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT DISTINCT categoria FROM produtos WHERE categoria IS NOT NULL AND categoria <> \'\' ORDER BY categoria;")\r\n'
    '   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("categoria")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optETags_Click()\r\n'
    'Dim r As ADODB.Recordset\r\n'
    '   lblEdicaoColetiva.Caption = "Tags"\r\n'
    '   txtEdicaoColetiva.Visible = False\r\n'
    '   cboEdicaoColetiva.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> \'\' ORDER BY TAGS;")\r\n'
    '   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("TAGS")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optECBS_Click()\r\n'
    'Dim r As ADODB.Recordset\r\n'
    '   lblEdicaoColetiva.Caption = "CBS Class."\r\n'
    '   txtEdicaoColetiva.Visible = False\r\n'
    '   cboEdicaoColetiva.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT cClassTrib + \' - \' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;")\r\n'
    '   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("val")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optEIS_Click()\r\n'
    'Dim r As ADODB.Recordset\r\n'
    '   lblEdicaoColetiva.Caption = "IS Class."\r\n'
    '   txtEdicaoColetiva.Visible = False\r\n'
    '   cboEdicaoColetiva.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT cClassTrib_IS + \' - \' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;")\r\n'
    '   Do While Not r.EOF: cboEdicaoColetiva.AddItem ValidateNull(r("val")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdicaoColetiva.Visible = True\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cmdEdicaoColetiva_Click()\r\n'
    'Dim i As Integer\r\n'
    'Dim sVal As String\r\n'
    'Dim iCol As Integer\r\n'
    'Dim sSet As String\r\n'
    'Dim p As Integer\r\n'
    '\r\n'
    'If txtEdicaoColetiva.Visible Then\r\n'
    '   sVal = txtEdicaoColetiva.Text\r\n'
    'Else\r\n'
    '   sVal = cboEdicaoColetiva.Text\r\n'
    'End If\r\n'
    '\r\n'
    'If optEICMS.Value Then\r\n'
    '   iCol = 10: sSet = "ICMSCST = \'" & sVal & "\'"\r\n'
    'ElseIf optENCM.Value Then\r\n'
    '   iCol = 8: sSet = "NCM = \'" & sVal & "\'"\r\n'
    'ElseIf optECFOP.Value Then\r\n'
    '   iCol = 9: sSet = "CFOP = " & sVal\r\n'
    'ElseIf optECest.Value Then\r\n'
    '   iCol = 18: sSet = "cest = \'" & sVal & "\'"\r\n'
    'ElseIf optECategoria.Value Then\r\n'
    '   iCol = 6: sSet = "categoria = \'" & sVal & "\'"\r\n'
    'ElseIf optETags.Value Then\r\n'
    '   iCol = 7: sSet = "TAGS = \'" & sVal & "\'"\r\n'
    'ElseIf optECBS.Value Then\r\n'
    '   p = InStr(sVal, " - "): If p > 0 Then sVal = Left(sVal, p - 1)\r\n'
    '   iCol = 19: sSet = "cClassTrib = \'" & sVal & "\'"\r\n'
    'ElseIf optEIS.Value Then\r\n'
    '   p = InStr(sVal, " - "): If p > 0 Then sVal = Left(sVal, p - 1)\r\n'
    '   iCol = 20: sSet = "cClassTrib_IS = \'" & sVal & "\'"\r\n'
    'Else\r\n'
    '   Exit Sub\r\n'
    'End If\r\n'
    '\r\n'
    'If Len(Trim(sSet)) = 0 Then Exit Sub\r\n'
    '\r\n'
    'picAguarde.Visible = True\r\n'
    'DoEvents\r\n'
    '\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   Grid.TextMatrix(i, iCol) = sVal\r\n'
    '   dbData.Execute "UPDATE produtos SET " & sSet & " WHERE codigo = " & Grid.TextMatrix(i, 2) & ";"\r\n'
    'Next i\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('Insert edicao coletiva handlers',
    'Grid.ColWidth(20) = 0\r\nEnd If\r\nEnd Sub\r\n\r\nPrivate Sub Form_Unload',
    'Grid.ColWidth(20) = 0\r\nEnd If\r\nEnd Sub\r\n\r\n' + new_handlers + 'Private Sub Form_Unload',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
