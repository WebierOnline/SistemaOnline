data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1: GridPedidos_MouseDown – usar RowData, sem TextMatrix col 0 ──────
old1 = (
    'Private Sub GridPedidos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)\r\n'
    'If x > GridPedidos.ColWidth(0) Then Exit Sub\r\n'
    '\r\n'
    'Dim iRow As Long, yPos As Single\r\n'
    'yPos = GridPedidos.RowHeight(0)\r\n'
    'If y < yPos Then Exit Sub\r\n'
    'For iRow = 1 To GridPedidos.rows - 1\r\n'
    '    yPos = yPos + GridPedidos.RowHeight(iRow)\r\n'
    '    If y < yPos Then\r\n'
    '        If GridPedidos.TextMatrix(iRow, 0) = "1" Then\r\n'
    '            GridPedidos.TextMatrix(iRow, 0) = ""\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '            GridPedidos.CellForeColor = GridPedidos.CellBackColor\r\n'
    '        Else\r\n'
    '            GridPedidos.TextMatrix(iRow, 0) = "1"\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '            GridPedidos.CellForeColor = GridPedidos.CellBackColor\r\n'
    '        End If\r\n'
    '        Exit Sub\r\n'
    '    End If\r\n'
    'Next iRow\r\n'
    'End Sub'
)
new1 = (
    'Private Sub GridPedidos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)\r\n'
    'If x > GridPedidos.ColWidth(0) Then Exit Sub\r\n'
    '\r\n'
    'Dim iRow As Long, yPos As Single\r\n'
    'yPos = GridPedidos.RowHeight(0)\r\n'
    'If y < yPos Then Exit Sub\r\n'
    'For iRow = 1 To GridPedidos.rows - 1\r\n'
    '    yPos = yPos + GridPedidos.RowHeight(iRow)\r\n'
    '    If y < yPos Then\r\n'
    '        If GridPedidos.RowData(iRow) = 1 Then\r\n'
    '            GridPedidos.RowData(iRow) = 0\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '        Else\r\n'
    '            GridPedidos.RowData(iRow) = 1\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '        End If\r\n'
    '        Exit Sub\r\n'
    '    End If\r\n'
    'Next iRow\r\n'
    'End Sub'
)

# ── Patch 2: imgDesmarcadaTODAS_Click – usar RowData ─────────────────────────
old2 = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    GridPedidos.TextMatrix(i, 0) = "1"\r\n'
    '    GridPedidos.Row = i: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    '    GridPedidos.CellForeColor = GridPedidos.CellBackColor\r\n'
    'Next i\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ImgMarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    GridPedidos.TextMatrix(i, 0) = ""\r\n'
    '    GridPedidos.Row = i: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    '    GridPedidos.CellForeColor = GridPedidos.CellBackColor\r\n'
    'Next i\r\n'
    'End Sub'
)
new2 = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    GridPedidos.RowData(i) = 1\r\n'
    '    GridPedidos.Row = i: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ImgMarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    GridPedidos.RowData(i) = 0\r\n'
    '    GridPedidos.Row = i: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'End Sub'
)

# ── Patch 3: cmdConverterNFe_Click – checar RowData em vez de TextMatrix ──────
old3a = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 0) = "1" And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        nMarcados = nMarcados + 1\r\n'
    '    End If\r\n'
    'Next i'
)
new3a = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.RowData(i) = 1 And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        nMarcados = nMarcados + 1\r\n'
    '    End If\r\n'
    'Next i'
)

old3b = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.TextMatrix(i, 0) = "1" And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        txtCodPedido.Text = GridPedidos.TextMatrix(i, 1)\r\n'
    '        vTipoEdicaoNFe = "Edicao"\r\n'
    '        GravarPedido\r\n'
    '        AtualizarTotaisNota\r\n'
    '        txtInfComple.Text = "NFe referente a venda N\xba " & txtCodPedido.Text\r\n'
    '    End If\r\n'
    'Next i'
)
new3b = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    If GridPedidos.RowData(i) = 1 And GridPedidos.TextMatrix(i, 2) <> "SIM" Then\r\n'
    '        txtCodPedido.Text = GridPedidos.TextMatrix(i, 1)\r\n'
    '        vTipoEdicaoNFe = "Edicao"\r\n'
    '        GravarPedido\r\n'
    '        AtualizarTotaisNota\r\n'
    '        txtInfComple.Text = "NFe referente a venda N\xba " & txtCodPedido.Text\r\n'
    '    End If\r\n'
    'Next i'
)

# ── Patch 4: FormatarGridPedidos – remover TextMatrix(0) col 0 e CellForeColor
old4 = (
    '            .Row = .rows - 1: .Col = 0\r\n'
    '            Set .CellPicture = imgDesmarcada.Picture\r\n'
    '            .CellPictureAlignment = 4\r\n'
    '            .CellForeColor = .CellBackColor'
)
new4 = (
    '            .Row = .rows - 1: .Col = 0\r\n'
    '            Set .CellPicture = imgDesmarcada.Picture\r\n'
    '            .CellPictureAlignment = 4'
)

checks = [(old1,new1,'MouseDown'),(old2,new2,'TODAS'),(old3a,new3a,'conv_count'),
          (old3b,new3b,'conv_loop'),(old4,new4,'init')]
all_ok = True
for old, new, name in checks:
    c = text.count(old)
    print(f'{name}: {c}')
    if c != 1: all_ok = False

if all_ok:
    for old, new, _ in checks:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n',b'\n').replace(b'\r',b'\n').replace(b'\n',b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm','wb').write(out)
    print('OK')
else:
    print('ABORTADO')
