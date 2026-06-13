data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

# ── Patch 1a: FormatarGridPedidos – ocultar texto da célula col 0 no init ───
old1a = (
    '            .Row = .rows - 1: .Col = 0\r\n'
    '            Set .CellPicture = imgDesmarcada.Picture\r\n'
    '            .CellPictureAlignment = 4'
)
new1a = (
    '            .Row = .rows - 1: .Col = 0\r\n'
    '            Set .CellPicture = imgDesmarcada.Picture\r\n'
    '            .CellPictureAlignment = 4\r\n'
    '            .CellForeColor = .CellBackColor'
)

# ── Patch 1b: GridPedidos_MouseDown – ocultar "1" após marcar/desmarcar ─────
old1b = (
    '        If GridPedidos.TextMatrix(iRow, 0) = "1" Then\r\n'
    '            GridPedidos.TextMatrix(iRow, 0) = ""\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '        Else\r\n'
    '            GridPedidos.TextMatrix(iRow, 0) = "1"\r\n'
    '            GridPedidos.Row = iRow: GridPedidos.Col = 0\r\n'
    '            Set GridPedidos.CellPicture = ImgMarcada.Picture\r\n'
    '            GridPedidos.CellPictureAlignment = 4\r\n'
    '        End If'
)
new1b = (
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
    '        End If'
)

# ── Patch 1c: imgDesmarcadaTODAS_Click – ocultar texto no loop ───────────────
old1c = (
    'For i = 1 To GridPedidos.rows - 1\r\n'
    '    GridPedidos.TextMatrix(i, 0) = "1"\r\n'
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
    '    GridPedidos.TextMatrix(i, 0) = ""\r\n'
    '    GridPedidos.Row = i: GridPedidos.Col = 0\r\n'
    '    Set GridPedidos.CellPicture = imgDesmarcada.Picture\r\n'
    '    GridPedidos.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'End Sub'
)
new1c = (
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

# ── Patch 2: cmdConverterNFe_Click – refresh grid nos dois casos ─────────────
old2 = (
    'If nMarcados = 1 Then\r\n'
    '    cmdNovo.Enabled = False\r\n'
    '    cmdSalvar.Enabled = True\r\n'
    '    cmdCancelar.Enabled = True\r\n'
    '    frmNota.Enabled = True\r\n'
    '    frmDestinatario.Enabled = True\r\n'
    '    frmItens.Enabled = True\r\n'
    '    Tab_Totais.Enabled = True\r\n'
    '    Tab_Produtos.Enabled = True\r\n'
    '    cmdRecalcular_Click\r\n'
    '    Frm_NF.Tab = 0\r\n'
    'Else\r\n'
    '    vTipoEdicaoNFe = ""\r\n'
    '    cmdNovo.Enabled = True\r\n'
    '    cmdSalvar.Enabled = False\r\n'
    '    cmdCancelar.Enabled = False\r\n'
    '    frmNota.Enabled = False\r\n'
    '    frmDestinatario.Enabled = False\r\n'
    '    frmItens.Enabled = False\r\n'
    '    Tab_Totais.Enabled = False\r\n'
    '    Tab_Produtos.Enabled = False\r\n'
    '    cmdExibirPedidos_Click\r\n'
    'End If'
)
new2 = (
    'cmdExibirPedidos_Click\r\n'
    '\r\n'
    'If nMarcados = 1 Then\r\n'
    '    cmdNovo.Enabled = False\r\n'
    '    cmdSalvar.Enabled = True\r\n'
    '    cmdCancelar.Enabled = True\r\n'
    '    frmNota.Enabled = True\r\n'
    '    frmDestinatario.Enabled = True\r\n'
    '    frmItens.Enabled = True\r\n'
    '    Tab_Totais.Enabled = True\r\n'
    '    Tab_Produtos.Enabled = True\r\n'
    '    cmdRecalcular_Click\r\n'
    '    Frm_NF.Tab = 0\r\n'
    'Else\r\n'
    '    vTipoEdicaoNFe = ""\r\n'
    '    cmdNovo.Enabled = True\r\n'
    '    cmdSalvar.Enabled = False\r\n'
    '    cmdCancelar.Enabled = False\r\n'
    '    frmNota.Enabled = False\r\n'
    '    frmDestinatario.Enabled = False\r\n'
    '    frmItens.Enabled = False\r\n'
    '    Tab_Totais.Enabled = False\r\n'
    '    Tab_Produtos.Enabled = False\r\n'
    'End If'
)

checks = [(old1a,new1a,'init_pic'),(old1b,new1b,'mousedown'),(old1c,new1c,'TODAS_loops'),(old2,new2,'converter')]
all_ok = True
for old, new, name in checks:
    c = text.count(old)
    print(f'{name}: count={c}')
    if c != 1:
        all_ok = False

if all_ok:
    for old, new, _ in checks:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK - 4 patches aplicados')
else:
    print('ABORTADO')
