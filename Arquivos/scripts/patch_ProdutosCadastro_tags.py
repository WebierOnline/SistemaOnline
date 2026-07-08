path = r'C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm'
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

# ── 1: cboTAGs_GotFocus — consultar Categorias_Tags ao invés de produtos ──
content = sub('cboTAGs query',
    'RsOpen rTags, "SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> \'\' AND categoria = \'" & Replace(cboCategoria.Text, "\'", "\'\'") & "\' ORDER BY TAGS"\r\nDo While Not rTags.EOF\r\n    cboTAGs.AddItem rTags("TAGS")\r\n    rTags.MoveNext\r\nLoop',
    'RsOpen rTags, "SELECT ct.Tags FROM Categorias_Tags ct INNER JOIN Categorias c ON ct.ID_Categoria = c.ID_Categoria WHERE c.Categoria = \'" & Replace(cboCategoria.Text, "\'", "\'\'") & "\' ORDER BY ct.Tags"\r\nDo While Not rTags.EOF\r\n    cboTAGs.AddItem rTags("Tags")\r\n    rTags.MoveNext\r\nLoop',
    content)

# ── 2: cmdSalvar_Click — capturar tag/categoria antes de qualquer Exit Sub ─
content = sub('cmdSalvar capture values',
    'Private Sub cmdSalvar_Click()\r\nIf txtEAN.Text <> "" Then',
    'Private Sub cmdSalvar_Click()\r\nDim sSaveTag As String\r\nDim sSaveCat As String\r\nsSaveTag = Trim(cboTAGs.Text)\r\nsSaveCat = Trim(cboCategoria.Text)\r\nIf txtEAN.Text <> "" Then',
    content)

# ── 3: cmdSalvar_Click — auto-inserir nova tag ao final do save ───────────
auto_tag = (
    'If Len(sSaveTag) > 0 And Len(sSaveCat) > 0 Then\r\n'
    '   Dim rAutoTag As ADODB.Recordset\r\n'
    '   Dim lCatID As Long\r\n'
    '   RsOpen rAutoTag, "SELECT ID_Categoria FROM Categorias WHERE Categoria = \'" & Replace(sSaveCat, "\'", "\'\'") & "\'"\r\n'
    '   If Not rAutoTag.EOF Then lCatID = CLng(rAutoTag("ID_Categoria"))\r\n'
    '   If rAutoTag.State <> 0 Then rAutoTag.Close\r\n'
    '   If lCatID > 0 Then\r\n'
    '      RsOpen rAutoTag, "SELECT COUNT(*) AS qtd FROM Categorias_Tags WHERE Tags = \'" & Replace(sSaveTag, "\'", "\'\'") & "\' AND ID_Categoria = " & lCatID\r\n'
    '      If Not rAutoTag.EOF Then\r\n'
    '         If CLng(rAutoTag("qtd")) = 0 Then\r\n'
    '            dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES (\'" & Replace(sSaveTag, "\'", "\'\'") & "\', " & lCatID & ");"\r\n'
    '         End If\r\n'
    '      End If\r\n'
    '      If rAutoTag.State <> 0 Then rAutoTag.Close\r\n'
    '   End If\r\nEnd If\r\n'
)

content = sub('cmdSalvar auto-tag',
    'SSTab2.Tab = 0\r\n    cmdExibir_Click\r\nEnd If\r\nEnd Sub',
    'SSTab2.Tab = 0\r\n    cmdExibir_Click\r\nEnd If\r\n' + auto_tag + 'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
