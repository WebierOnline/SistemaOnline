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
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: cmdEdicaoColetiva_Click — auto-insert tag nova na Categorias_Tags ──
# Apos o loop principal, antes de picAguarde.Visible = False
content = sub('cmdEdicaoColetiva auto-insert tag',
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub',

    'If optETags.Value And Len(Trim(sVal)) > 0 Then\r\n'
    '   Dim rCat As ADODB.Recordset\r\n'
    '   Dim sCatNome As String\r\n'
    '   Dim lCatID As Long\r\n'
    '   Dim j As Integer\r\n'
    '   For j = 1 To Grid.rows - 1\r\n'
    '      If bTemMarcadas And Grid.TextMatrix(j, 0) <> "1" Then GoTo ProxCat\r\n'
    '      sCatNome = Grid.TextMatrix(j, 6)\r\n'
    '      If Len(Trim(sCatNome)) = 0 Then GoTo ProxCat\r\n'
    '      lCatID = 0\r\n'
    '      Set rCat = dbData.OpenRecordset("SELECT ID_Categoria FROM Categorias WHERE Categoria = \'" & Replace(sCatNome, "\'", "\'\'") & "\'")\r\n'
    '      If Not rCat.EOF Then lCatID = CLng(rCat("ID_Categoria"))\r\n'
    '      If rCat.State <> 0 Then rCat.Close\r\n'
    '      If lCatID > 0 Then\r\n'
    '         Set rCat = dbData.OpenRecordset("SELECT COUNT(*) AS qtd FROM Categorias_Tags WHERE Tags = \'" & Replace(sVal, "\'", "\'\'") & "\' AND ID_Categoria = " & lCatID)\r\n'
    '         If Not rCat.EOF Then\r\n'
    '            If CLng(rCat("qtd")) = 0 Then\r\n'
    '               dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES (\'" & Replace(sVal, "\'", "\'\'") & "\', " & lCatID & ");"\r\n'
    '            End If\r\n'
    '         End If\r\n'
    '         If rCat.State <> 0 Then rCat.Close\r\n'
    '      End If\r\n'
    '      ProxCat:\r\n'
    '   Next j\r\n'
    'End If\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub',
    content)

# ── 2: Novos handlers cboEdicaoColetiva — somente letras maiusculas ────────
new_handlers = (
    'Private Sub cboEdicaoColetiva_KeyPress(KeyAscii As Integer)\r\n'
    '   If Not optETags.Value Then Exit Sub\r\n'
    '   If KeyAscii >= 97 And KeyAscii <= 122 Then\r\n'
    '      KeyAscii = KeyAscii - 32\r\n'
    '   ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8) Then\r\n'
    '      KeyAscii = 0\r\n'
    '   End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdicaoColetiva_Change()\r\n'
    '   If Not optETags.Value Then Exit Sub\r\n'
    '   Dim pos As Integer\r\n'
    '   pos = cboEdicaoColetiva.SelStart\r\n'
    '   cboEdicaoColetiva.Text = UCase(cboEdicaoColetiva.Text)\r\n'
    '   cboEdicaoColetiva.SelStart = pos\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('Insert cboEdicaoColetiva handlers',
    'Private Sub ResetarMarcas()',
    new_handlers + 'Private Sub ResetarMarcas()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
