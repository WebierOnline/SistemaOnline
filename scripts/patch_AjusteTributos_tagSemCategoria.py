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

# ── 1: Grid_Click Case 7 — bloquear edicao de tag sem categoria ───────────
content = sub('Case 7 bloquear sem categoria',
    'Case 7\r\n'
    '   cboEdit.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT ct.Tags',

    'Case 7\r\n'
    '   If Len(Trim(Grid.TextMatrix(iRow, 6))) = 0 Then\r\n'
    '      MsgBox "Defina a categoria do produto antes de editar a tag.", vbExclamation, "Tag sem categoria"\r\n'
    '      Exit Sub\r\n'
    '   End If\r\n'
    '   cboEdit.Clear\r\n'
    '   Set r = dbData.OpenRecordset("SELECT ct.Tags',
    content)

# ── 2: cmdEdicaoColetiva — validar linhas sem categoria antes de salvar tags
validacao_cat = (
    'If optETags.Value Then\r\n'
    '   Dim iLinSemCat As Integer\r\n'
    '   For i = 1 To Grid.rows - 1\r\n'
    '      If bTemMarcadas And Grid.TextMatrix(i, 0) <> "1" Then GoTo ProxValCat\r\n'
    '      If Len(Trim(Grid.TextMatrix(i, 6))) = 0 Then iLinSemCat = iLinSemCat + 1\r\n'
    '      ProxValCat:\r\n'
    '   Next i\r\n'
    '   If iLinSemCat > 0 Then\r\n'
    '      MsgBox iLinSemCat & " produto(s) sem categoria. Defina a categoria antes de aplicar a tag.", vbExclamation, "Tag sem categoria"\r\n'
    '      Exit Sub\r\n'
    '   End If\r\n'
    'End If\r\n'
    '\r\n'
)

content = sub('cmdEdicaoColetiva validar sem categoria',
    'For i = 1 To Grid.rows - 1\r\n'
    '   If Grid.TextMatrix(i, 0) = "1" Then bTemMarcadas = True: Exit For\r\n'
    'Next i\r\n'
    '\r\n'
    'picAguarde.Visible = True',

    'For i = 1 To Grid.rows - 1\r\n'
    '   If Grid.TextMatrix(i, 0) = "1" Then bTemMarcadas = True: Exit For\r\n'
    'Next i\r\n'
    '\r\n'
    + validacao_cat +
    'picAguarde.Visible = True',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
