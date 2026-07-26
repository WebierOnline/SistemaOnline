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

# ── 1: Formatar_Grid — centralizar imagem + reset estado "marcar todos" ───
content = sub('Formatar_Grid picAlign + reset',
    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '   Next lChk\r\n'
    '   End With\r\n'
    '   chkPISCOFINS_Click\r\n'
    '   chkIPI_Click\r\n'
    '   chkCest_Click\r\n'
    '   chkCBSIS_Click\r\n'
    'End Sub',

    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '      Grid.CellPictureAlignment = 4\r\n'
    '   Next lChk\r\n'
    '   End With\r\n'
    '   chkPISCOFINS_Click\r\n'
    '   chkIPI_Click\r\n'
    '   chkCest_Click\r\n'
    '   chkCBSIS_Click\r\n'
    '   imgDesmarcadaTODAS.Visible = True\r\n'
    '   ImgMarcadaTODAS.Visible = False\r\n'
    '   lblMarcarTodas.Caption = "Marcar todos"\r\n'
    '   AvaliarFrmAlterarGrupos\r\n'
    'End Sub',
    content)

# ── 2: Grid_Click Case 0 — fix click + centralizar + chamar avaliacao ─────
content = sub('Grid_Click Case 0 fix',
    'Case 0\r\n'
    '   Grid.Row = iRow\r\n'
    '   Grid.Col = 0\r\n'
    '   If Grid.TextMatrix(iRow, 0) = "1" Then\r\n'
    '      Grid.TextMatrix(iRow, 0) = ""\r\n'
    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '   Else\r\n'
    '      Grid.TextMatrix(iRow, 0) = "1"\r\n'
    '      Set Grid.CellPicture = ImgMarcada.Picture\r\n'
    '   End If',

    'Case 0\r\n'
    '   If iRow < 1 Then Exit Sub\r\n'
    '   If Grid.TextMatrix(iRow, 0) = "1" Then\r\n'
    '      Grid.TextMatrix(iRow, 0) = ""\r\n'
    '      Grid.Row = iRow: Grid.Col = 0\r\n'
    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '      Grid.CellPictureAlignment = 4\r\n'
    '   Else\r\n'
    '      Grid.TextMatrix(iRow, 0) = "1"\r\n'
    '      Grid.Row = iRow: Grid.Col = 0\r\n'
    '      Set Grid.CellPicture = ImgMarcada.Picture\r\n'
    '      Grid.CellPictureAlignment = 4\r\n'
    '   End If\r\n'
    '   AvaliarFrmAlterarGrupos',
    content)

# ── 3: Inserir handlers + helper antes de Form_Unload ─────────────────────
new_handlers = (
    'Private Sub AvaliarFrmAlterarGrupos()\r\n'
    'Dim j As Integer\r\n'
    'Dim bEnabled As Boolean\r\n'
    'For j = 1 To Grid.rows - 1\r\n'
    '   If Grid.TextMatrix(j, 0) = "1" Then bEnabled = True: Exit For\r\n'
    'Next j\r\n'
    'If Not bEnabled Then\r\n'
    '   If Not (optTodos.Value Or optCodBarra.Value Or optDesc.Value) Then\r\n'
    '      If Grid.rows > 1 Then bEnabled = True\r\n'
    '   End If\r\n'
    'End If\r\n'
    'frmAlterarGrupos.Enabled = bEnabled\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub imgDesmarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'imgDesmarcadaTODAS.Visible = False\r\n'
    'ImgMarcadaTODAS.Visible = True\r\n'
    'lblMarcarTodas.Caption = "Desmarcar todos"\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   Grid.TextMatrix(i, 0) = "1"\r\n'
    '   Grid.Row = i: Grid.Col = 0\r\n'
    '   Set Grid.CellPicture = ImgMarcada.Picture\r\n'
    '   Grid.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'AvaliarFrmAlterarGrupos\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub ImgMarcadaTODAS_Click()\r\n'
    'Dim i As Integer\r\n'
    'ImgMarcadaTODAS.Visible = False\r\n'
    'imgDesmarcadaTODAS.Visible = True\r\n'
    'lblMarcarTodas.Caption = "Marcar todos"\r\n'
    'For i = 1 To Grid.rows - 1\r\n'
    '   Grid.TextMatrix(i, 0) = ""\r\n'
    '   Grid.Row = i: Grid.Col = 0\r\n'
    '   Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '   Grid.CellPictureAlignment = 4\r\n'
    'Next i\r\n'
    'AvaliarFrmAlterarGrupos\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('Insert marcaTodas handlers',
    'picAguarde.Visible = False\r\nEnd Sub\r\n\r\nPrivate Sub Form_Unload',
    'picAguarde.Visible = False\r\nEnd Sub\r\n\r\n' + new_handlers + 'Private Sub Form_Unload',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
