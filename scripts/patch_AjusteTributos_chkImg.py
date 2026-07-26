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

# ── 1: Formatar_Grid — dar largura para col 0 ─────────────────────────────
content = sub('ColWidth(0) largura',
    '      .Cols = 21\r\n'
    '      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 0',

    '      .Cols = 21\r\n'
    '      .rows = 2\r\n'
    '      \r\n'
    '      .ColWidth(0) = 360',
    content)

# ── 2: Formatar_Grid — inicializar imgDesmarcada em cada linha ────────────
content = sub('Formatar_Grid init checkbox pics',
    'picAguarde.Visible = False\r\n'
    '   End With\r\n'
    '   chkPISCOFINS_Click',

    'picAguarde.Visible = False\r\n'
    '   Dim lChk As Long\r\n'
    '   For lChk = 1 To Grid.rows - 1\r\n'
    '      Grid.Row = lChk\r\n'
    '      Grid.Col = 0\r\n'
    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '   Next lChk\r\n'
    '   End With\r\n'
    '   chkPISCOFINS_Click',
    content)

# ── 3: Grid_Click — Case 0 para alternar marcada/desmarcada ──────────────
content = sub('Grid_Click Case 0 toggle',
    'Select Case iCol\r\n'
    'Case 6',

    'Select Case iCol\r\n'
    'Case 0\r\n'
    '   Grid.Row = iRow\r\n'
    '   Grid.Col = 0\r\n'
    '   If Grid.TextMatrix(iRow, 0) = "1" Then\r\n'
    '      Grid.TextMatrix(iRow, 0) = ""\r\n'
    '      Set Grid.CellPicture = imgDesmarcada.Picture\r\n'
    '   Else\r\n'
    '      Grid.TextMatrix(iRow, 0) = "1"\r\n'
    '      Set Grid.CellPicture = imgMarcada.Picture\r\n'
    '   End If\r\n'
    'Case 6',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
