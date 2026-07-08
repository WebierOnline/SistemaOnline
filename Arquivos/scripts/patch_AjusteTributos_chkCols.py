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
    return c.replace(old, new, 1)

# ── 1: Formatar_Grid — chamar handlers apos End With ───────────────────────
content = sub('Formatar_Grid sync chk',
    '      .rows = .rows - 1\r\n'
    '      .Redraw = True\r\n'
    '      picAguarde.Visible = False\r\n'
    '   End With\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub For',      # "Form_Load" ou "Form_Unload" — qualquer que venha depois

    '      .rows = .rows - 1\r\n'
    '      .Redraw = True\r\n'
    '      picAguarde.Visible = False\r\n'
    '   End With\r\n'
    '   chkPISCOFINS_Click\r\n'
    '   chkIPI_Click\r\n'
    '   chkCest_Click\r\n'
    '   chkCBSIS_Click\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub For',
    content)

if 'Formatar_Grid sync chk' in [e.split(':')[0] for e in errors]:
    print('ERRO no ponto 1 — tentando variante sem prefixo "For"')
    errors.clear()

# ── 2: Handlers dos checkboxes — inserir antes de Form_Unload ──────────────
new_handlers = (
    'Private Sub chkPISCOFINS_Click()\r\n'
    'If chkPISCOFINS.Value = Checked Then\r\n'
    '   Grid.ColWidth(12) = 650\r\n'
    '   Grid.ColWidth(13) = 700\r\n'
    '   Grid.ColWidth(14) = 800\r\n'
    '   Grid.ColWidth(15) = 700\r\n'
    'Else\r\n'
    '   Grid.ColWidth(12) = 0\r\n'
    '   Grid.ColWidth(13) = 0\r\n'
    '   Grid.ColWidth(14) = 0\r\n'
    '   Grid.ColWidth(15) = 0\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub chkIPI_Click()\r\n'
    'If chkIPI.Value = Checked Then\r\n'
    '   Grid.ColWidth(16) = 650\r\n'
    '   Grid.ColWidth(17) = 700\r\n'
    'Else\r\n'
    '   Grid.ColWidth(16) = 0\r\n'
    '   Grid.ColWidth(17) = 0\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub chkCest_Click()\r\n'
    'If chkCest.Value = Checked Then\r\n'
    '   Grid.ColWidth(18) = 900\r\n'
    'Else\r\n'
    '   Grid.ColWidth(18) = 0\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
    'Private Sub chkCBSIS_Click()\r\n'
    'If chkCBSIS.Value = Checked Then\r\n'
    '   Grid.ColWidth(19) = 900\r\n'
    '   Grid.ColWidth(20) = 900\r\n'
    'Else\r\n'
    '   Grid.ColWidth(19) = 0\r\n'
    '   Grid.ColWidth(20) = 0\r\n'
    'End If\r\n'
    'End Sub\r\n'
    '\r\n'
)

content = sub('chk handlers',
    'Private Sub Form_Unload(Cancel As Integer)\r\n',
    new_handlers + 'Private Sub Form_Unload(Cancel As Integer)\r\n',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK — arquivo gravado.')
