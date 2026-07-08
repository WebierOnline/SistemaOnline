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

# ── 1: Adicionar cboEdit no layout antes do End que fecha o Form ────────────
cbo_def = (
    '   Begin VB.ComboBox cboEdit \r\n'
    '      Height          =   315\r\n'
    '      Left            =   120\r\n'
    '      TabIndex        =   42\r\n'
    '      Top             =   120\r\n'
    '      Visible         =   0   \'False\r\n'
    '      Width           =   1200\r\n'
    '   End\r\n'
    'End\r\n'
    'Attribute VB_Name'
)
content = sub('cboEdit layout',
    'End\r\nAttribute VB_Name',
    cbo_def,
    content)

# ── 2: Substituir Grid_Click por Select Case com cboEdit ───────────────────
old_grid_click = (
    'Private Sub Grid_Click()\r\n'
    'Dim i As Integer\r\n'
    '\r\n'
    'For i = 6 To 20\r\n'
    '   If Grid.ColSel = i Then\r\n'
    '      txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '      txtEdit.Text = Grid.TextMatrix(Grid.Row, Grid.Col)\r\n'
    '      txtEdit.Visible = True\r\n'
    '      txtEdit.SetFocus\r\n'
    '      txtEdit.SelStart = 0\r\n'
    '      txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    '      iRow = Grid.Row\r\n'
    '      iCol = Grid.Col\r\n'
    '   End If\r\n'
    'Next\r\n'
    'End Sub\r\n'
)

new_grid_click = (
    'Private Sub Grid_Click()\r\n'
    'Dim r As ADODB.Recordset\r\n'
    '\r\n'
    'txtEdit.Visible = False\r\n'
    'cboEdit.Visible = False\r\n'
    'iRow = Grid.Row\r\n'
    'iCol = Grid.Col\r\n'
    '\r\n'
    'Select Case iCol\r\n'
    'Case 6\r\n'
    "   cboEdit.Clear\r\n"
    "   Set r = dbData.OpenRecordset(\"SELECT DISTINCT categoria FROM produtos WHERE categoria IS NOT NULL AND categoria <> '' ORDER BY categoria;\")\r\n"
    '   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("categoria")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '   cboEdit.Visible = True\r\n'
    '   cboEdit.SetFocus\r\n'
    'Case 7\r\n'
    '   cboEdit.Clear\r\n'
    "   Set r = dbData.OpenRecordset(\"SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' ORDER BY TAGS;\")\r\n"
    '   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("TAGS")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '   cboEdit.Visible = True\r\n'
    '   cboEdit.SetFocus\r\n'
    'Case 19\r\n'
    '   cboEdit.Clear\r\n'
    "   Set r = dbData.OpenRecordset(\"SELECT cClassTrib + ' - ' + NomecClassTrib AS val FROM TbIBSCBSClassTrib GROUP BY cClassTrib, NomecClassTrib ORDER BY cClassTrib;\")\r\n"
    '   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("val")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '   cboEdit.Visible = True\r\n'
    '   cboEdit.SetFocus\r\n'
    'Case 20\r\n'
    '   cboEdit.Clear\r\n'
    "   Set r = dbData.OpenRecordset(\"SELECT cClassTrib_IS + ' - ' + Descricao AS val FROM tbISClassTrib GROUP BY cClassTrib_IS, Descricao ORDER BY cClassTrib_IS;\")\r\n"
    '   Do While Not r.EOF: cboEdit.AddItem ValidateNull(r("val")): r.MoveNext: Loop\r\n'
    '   r.Close: Set r = Nothing\r\n'
    '   cboEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '   cboEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '   cboEdit.Visible = True\r\n'
    '   cboEdit.SetFocus\r\n'
    'Case 8 To 18\r\n'
    '   txtEdit.Move Grid.Left + Grid.CellLeft, Grid.Top + Grid.CellTop, Grid.CellWidth, Grid.CellHeight\r\n'
    '   txtEdit.Text = Grid.TextMatrix(iRow, iCol)\r\n'
    '   txtEdit.Visible = True\r\n'
    '   txtEdit.SetFocus\r\n'
    '   txtEdit.SelStart = 0\r\n'
    '   txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    'End Select\r\n'
    'End Sub\r\n'
)

content = sub('Grid_Click Select Case', old_grid_click, new_grid_click, content)

# ── 3: Inserir handlers cboEdit_Click / LostFocus / SaveCboEdit ────────────
new_handlers = (
    '\r\n'
    'Private Sub cboEdit_Click()\r\n'
    '   SaveCboEdit\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub cboEdit_LostFocus()\r\n'
    '   SaveCboEdit\r\n'
    '   cboEdit.Visible = False\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub SaveCboEdit()\r\n'
    'Dim sVal As String\r\n'
    'Dim p As Integer\r\n'
    'sVal = cboEdit.Text\r\n'
    'If iCol = 19 Or iCol = 20 Then\r\n'
    '   p = InStr(sVal, " - ")\r\n'
    '   If p > 0 Then sVal = Left(sVal, p - 1)\r\n'
    'End If\r\n'
    'Grid.TextMatrix(iRow, iCol) = sVal\r\n'
    'End Sub\r\n'
)

# Inserir logo apos o End Sub do Grid_Click
content = sub('Insert cboEdit handlers',
    'End Select\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub optCategoria_Click()',

    'End Select\r\n'
    'End Sub\r\n'
    + new_handlers +
    '\r\n'
    'Private Sub optCategoria_Click()',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('File written successfully.')
