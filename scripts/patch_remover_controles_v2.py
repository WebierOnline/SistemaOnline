data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')
lines = text.split('\r\n')

# ── Localiza início e fim de um bloco Begin...End pelo nome do controle ──────
def find_block(lines, ctrl_name):
    for i, l in enumerate(lines):
        s = l.strip()
        if s.startswith('Begin ') and s.endswith(ctrl_name):
            depth = 0
            for j in range(i, len(lines)):
                t = lines[j].strip()
                if t.startswith('BeginProperty'): depth += 1
                elif t == 'EndProperty':           depth -= 1
                elif t.startswith('Begin '):       depth += 1
                elif t == 'End':
                    depth -= 1
                    if depth == 0:
                        return i, j   # inclusive
    return None, None

# ── Localiza início e fim de uma Sub/Private Sub pelo cabeçalho ──────────────
def find_sub(lines, header):
    for i, l in enumerate(lines):
        if l.strip() == header or l.strip() == 'Private ' + header:
            for j in range(i, len(lines)):
                if lines[j].strip() == 'End Sub':
                    return i, j
    return None, None

ctrl_names = [
    'mskInicialPedidos', 'mskFinalPedidos',
    'cmdCalPedidos1',    'cmdCalPedidos2',
    'txtConCodPedido',   'txtCodClientePedidos',
    'lblConsCodPedido',  'lblInicialPedidos', 'lblFinalPedidos',
]

sub_headers = [
    'Sub cmdCalPedidos1_Click()',
    'Sub cmdCalPedidos2_Click()',
    'Sub mskInicialPedidos_GotFocus()',
    'Sub mskInicialPedidos_KeyPress(KeyAscii As Integer)',
    'Sub mskInicialPedidos_LostFocus()',
    'Sub mskFinalPedidos_GotFocus()',
    'Sub mskFinalPedidos_KeyPress(KeyAscii As Integer)',
    'Sub mskFinalPedidos_LostFocus()',
]

ranges_to_remove = []   # lista de (start, end) inclusive

for name in ctrl_names:
    s, e = find_block(lines, name)
    if s is None:
        print(f'NAOACHEI ctrl: {name}')
    else:
        print(f'ctrl {name}: {s}-{e}')
        ranges_to_remove.append((s, e))

for hdr in sub_headers:
    s, e = find_sub(lines, hdr)
    if s is None:
        print(f'NAOACHEI sub: {hdr}')
    else:
        print(f'sub {hdr}: {s}-{e}')
        ranges_to_remove.append((s, e))

print(f'\nTotal de intervalos: {len(ranges_to_remove)}')

if len(ranges_to_remove) != len(ctrl_names) + len(sub_headers):
    print('ABORTADO - algum trecho não encontrado')
else:
    # Remove em ordem reversa para preservar índices
    ranges_to_remove.sort(key=lambda r: r[0], reverse=True)
    new_lines = list(lines)
    for s, e in ranges_to_remove:
        del new_lines[s:e+1]

    # ── Remover referências a txtCodClientePedidos em cboClientePedidos_LostFocus ──
    result = '\r\n'.join(new_lines)

    old_block = (
        'If cboClientePedidos.Text = "" Then txtCodClientePedidos.Text = "": Exit Sub\r\n'
        'If cboClientePedidos.ListIndex = -1 Then txtCodClientePedidos.Text = "": Exit Sub\r\n'
        '\r\n'
        'txtCodClientePedidos = cboClientePedidos.ItemData(cboClientePedidos.ListIndex)\r\n'
        '\r\n'
    )
    c_block = result.count(old_block)
    print(f'lostfocus block count: {c_block}')

    # ── Remover .Visible = False sobrando em cboIndicePedidos_Click ──────────
    old_vis = 'txtCodClientePedidos.Visible = False\r\n'
    c_vis = result.count(old_vis)
    print(f'txtCodClientePedidos.Visible count: {c_vis}')

    if c_block == 1 and c_vis == 1:
        result = result.replace(old_block, '')
        result = result.replace(old_vis, '')

        out = result.encode('windows-1252')
        out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
        open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
        print('OK - remoção completa')
    else:
        print(f'ERRO contagens: block={c_block} vis={c_vis}')
