data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1: substituir GridPedidos_Click por GridPedidos_MouseDown ──────────
old1 = (
    'Private Sub GridPedidos_Click()\r\n'
    'Dim iRow As Integer\r\n'
    'iRow = GridPedidos.Row\r\n'
    'If iRow < 1 Then Exit Sub\r\n'
    'If GridPedidos.Col <> 0 Then Exit Sub\r\n'
    'If GridPedidos.TextMatrix(iRow, 0) = "1" Then\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = ""\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Else\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = "1"\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'End If\r\n'
    'End Sub'
)
new1 = (
    'Private Sub GridPedidos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)\r\n'
    'Dim iRow As Integer, iCol As Integer\r\n'
    'iRow = GridPedidos.RowContaining(y)\r\n'
    'iCol = GridPedidos.ColContaining(x)\r\n'
    'If iRow < 1 Then Exit Sub\r\n'
    'If iCol <> 0 Then Exit Sub\r\n'
    'If GridPedidos.TextMatrix(iRow, 0) = "1" Then\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = ""\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Else\r\n'
    '    GridPedidos.TextMatrix(iRow, 0) = "1"\r\n'
    '    GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'End If\r\n'
    'End Sub'
)

# ── Patch 2: FormatarGridPedidos – exibir/ocultar controles conforme registros ──
old2 = (
    '      .rows = .rows - 1\r\n'
    '      .Redraw = True\r\n'
    '   End With\r\n'
    '   \r\n'
    '\'   lblValor.Caption = Format(SomaGrid(GridPedidos, 5), ocMONEY)\r\n'
    'End Sub'
)
new2 = (
    '      .rows = .rows - 1\r\n'
    '      .Redraw = True\r\n'
    '   End With\r\n'
    '   \r\n'
    '   If GridPedidos.rows > 1 Then\r\n'
    '       imgDesmarcadaTODAS.Visible = True\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '       lblCodFabrica(9).Visible = True\r\n'
    '   Else\r\n'
    '       imgDesmarcadaTODAS.Visible = False\r\n'
    '       ImgMarcadaTODAS.Visible = False\r\n'
    '       lblCodFabrica(9).Visible = False\r\n'
    '   End If\r\n'
    '   \r\n'
    '\'   lblValor.Caption = Format(SomaGrid(GridPedidos, 5), ocMONEY)\r\n'
    'End Sub'
)

c1 = text.count(old1)
c2 = text.count(old2)
print(f'Patch1:{c1}  Patch2:{c2}')

if c1 == 1 and c2 == 1:
    text = text.replace(old1, new1)
    text = text.replace(old2, new2)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK')
else:
    print('ERRO - verificar contagens')
