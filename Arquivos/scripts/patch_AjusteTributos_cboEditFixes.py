path = r'C:\Projeto\OnlineCommerce\Forms\Produtos_AjusteTributos.frm'
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

# ── 1: cboEdit_KeyPress e _Change para TAGS (iCol=7) uppercase ────────────
new_cboEdit_handlers = (
    'Private Sub cboEdit_KeyPress(KeyAscii As Integer)\r\n'
    '   If iCol <> 7 Then Exit Sub\r\n'
    '   If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdit_Change()\r\n'
    '   If iCol <> 7 Then Exit Sub\r\n'
    '   Dim pos As Integer\r\n'
    '   pos = cboEdit.SelStart\r\n'
    '   cboEdit.Text = UCase(cboEdit.Text)\r\n'
    '   cboEdit.SelStart = pos\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('cboEdit KeyPress+Change TAGS',
    'Private Sub cboEdit_Click()',
    new_cboEdit_handlers + 'Private Sub cboEdit_Click()',
    content)

# ── 2: cmdAtualizar — auto-insert nova tag em Categorias_Tags ─────────────
tag_insert_loop = (
    '\r\n'
    'Dim rTag As ADODB.Recordset\r\n'
    'Dim sCatNomeTmp As String\r\n'
    'Dim lCatIDTmp As Long\r\n'
    'Dim sTagTmp As String\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   sTagTmp = Trim(Grid.TextMatrix(i, 7))\r\n'
    '   sCatNomeTmp = Trim(Grid.TextMatrix(i, 6))\r\n'
    '   If Len(sTagTmp) > 0 And Len(sCatNomeTmp) > 0 Then\r\n'
    '      lCatIDTmp = 0\r\n'
    '      Set rTag = dbData.OpenRecordset("SELECT ID_Categoria FROM Categorias WHERE Categoria = \'" & Replace(sCatNomeTmp, "\'", "\'\'") & "\'")\r\n'
    '      If Not rTag.EOF Then lCatIDTmp = CLng(rTag("ID_Categoria"))\r\n'
    '      If rTag.State <> 0 Then rTag.Close\r\n'
    '      If lCatIDTmp > 0 Then\r\n'
    '         Set rTag = dbData.OpenRecordset("SELECT COUNT(*) AS qtd FROM Categorias_Tags WHERE Tags = \'" & Replace(sTagTmp, "\'", "\'\'") & "\' AND ID_Categoria = " & lCatIDTmp)\r\n'
    '         If Not rTag.EOF Then\r\n'
    '            If CLng(rTag("qtd")) = 0 Then\r\n'
    '               dbData.Execute "INSERT INTO Categorias_Tags (Tags, ID_Categoria) VALUES (\'" & Replace(sTagTmp, "\'", "\'\'") & "\', " & lCatIDTmp & ");"\r\n'
    '            End If\r\n'
    '         End If\r\n'
    '         If rTag.State <> 0 Then rTag.Close\r\n'
    '      End If\r\n'
    '   End If\r\n'
    'Next i\r\n'
)

content = sub('cmdAtualizar auto-insert Categorias_Tags',
    '   dbData.Execute sSQL\r\n'
    'Next\r\n'
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cmdConsultaNCMean_Click()',

    '   dbData.Execute sSQL\r\n'
    'Next\r\n'
    + tag_insert_loop +
    '\r\n'
    'picAguarde.Visible = False\r\n'
    'ResetarMarcas\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cmdConsultaNCMean_Click()',
    content)

# ── 3: Ocultar cboEdit/txtEdit ao mudar filtro ────────────────────────────
HIDE = 'cboEdit.Visible = False\r\ntxtEdit.Visible = False\r\n'

content = sub('optCategoria hide cboEdit',
    'Private Sub optCategoria_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Categoria"',
    'Private Sub optCategoria_Click()\r\n' + HIDE +
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Categoria"',
    content)

content = sub('optCodBarra hide cboEdit',
    'Private Sub optCodBarra_Click()\r\n'
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False',
    'Private Sub optCodBarra_Click()\r\n' + HIDE +
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False',
    content)

content = sub('optDesc hide cboEdit',
    'Private Sub optDesc_Click()\r\n'
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False\r\n'
    'lblDesc.Visible = True',
    'Private Sub optDesc_Click()\r\n' + HIDE +
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False\r\n'
    'lblDesc.Visible = True',
    content)

content = sub('optTodos hide cboEdit',
    'Private Sub optTodos_Click()\r\n'
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False\r\n'
    'lblDesc.Visible = False',
    'Private Sub optTodos_Click()\r\n' + HIDE +
    'lblCategoria.Visible = False\r\n'
    'cboConsLinha.Visible = False\r\n'
    'lblDesc.Visible = False',
    content)

content = sub('optTags hide cboEdit',
    'Private Sub optTags_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Tags"',
    'Private Sub optTags_Click()\r\n' + HIDE +
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Tags"',
    content)

content = sub('optNCM hide cboEdit',
    'Private Sub optNCM_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "NCM"',
    'Private Sub optNCM_Click()\r\n' + HIDE +
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "NCM"',
    content)

content = sub('optClassTribCBS hide cboEdit',
    'Private Sub optClassTribCBS_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. CBS"',
    'Private Sub optClassTribCBS_Click()\r\n' + HIDE +
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. CBS"',
    content)

content = sub('optClassTribIS hide cboEdit',
    'Private Sub optClassTribIS_Click()\r\n'
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. IS"',
    'Private Sub optClassTribIS_Click()\r\n' + HIDE +
    'lblCategoria.Visible = True\r\n'
    'lblCategoria.Caption = "Classif. IS"',
    content)

content = sub('cmdLocalizar hide cboEdit',
    'Private Sub cmdLocalizar_Click()\r\n'
    'Dim sSQL As String\r\n'
    'Dim r As ADODB.Recordset',
    'Private Sub cmdLocalizar_Click()\r\n' + HIDE +
    'Dim sSQL As String\r\n'
    'Dim r As ADODB.Recordset',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
