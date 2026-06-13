data = open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb').read()
text = data.decode('windows-1252')

removals = []

# ── Utilitário: encontra bloco Begin..End completo ──────────────────────────
def find_begin_end(t, name):
    """Retorna (start, end) do bloco Begin...End que declara 'name'."""
    search = 'Begin ' + name
    idx = t.find(search)
    if idx == -1:
        return None, None
    # voltar ao início da linha
    start = t.rfind('\r\n', 0, idx) + 2
    # avançar até o End que fecha este bloco (depth=1)
    pos = idx
    depth = 0
    while pos < len(t):
        if t[pos:].startswith('Begin'):
            depth += 1
        if t[pos:pos+5] == '\r\nEnd' or (pos == 0 and t[:3] == 'End'):
            ln_start = pos + 2
            depth -= 1
            if depth == 0:
                end = t.find('\r\n', ln_start) + 2
                return start, end
        pos += 1
    return None, None

# ── Blocos de controle a remover ────────────────────────────────────────────
ctrl_names = [
    'MSMask.MaskEdBox mskInicialPedidos',
    'MSMask.MaskEdBox mskFinalPedidos',
    'ChamaleonBtn.chameleonButton cmdCalPedidos1',
    'ChamaleonBtn.chameleonButton cmdCalPedidos2',
    'VB.TextBox txtConCodPedido',
    'VB.TextBox txtCodClientePedidos',
    'VB.Label lblConsCodPedido',
    'VB.Label lblInicialPedidos',
    'VB.Label lblFinalPedidos',
]

# Marca cada bloco para remoção (do início da linha até após o End)
for name in ctrl_names:
    s, e = find_begin_end(text, name)
    if s is None:
        print(f'NAOACHEI: {name}')
    else:
        seg = text[s:e]
        c = text.count(seg)
        print(f'ctrl {name.split()[-1]}: count={c}')
        removals.append(seg)

# ── Subs de evento a remover ────────────────────────────────────────────────
subs_to_remove = [
    'Private Sub cmdCalPedidos1_Click()',
    'Private Sub cmdCalPedidos2_Click()',
    'Private Sub mskInicialPedidos_GotFocus()',
    'Private Sub mskInicialPedidos_KeyPress(KeyAscii As Integer)',
    'Private Sub mskInicialPedidos_LostFocus()',
    'Private Sub mskFinalPedidos_GotFocus()',
    'Private Sub mskFinalPedidos_KeyPress(KeyAscii As Integer)',
    'Private Sub mskFinalPedidos_LostFocus()',
]

def extract_sub(t, sub_header):
    idx = t.find(sub_header)
    if idx == -1:
        return None
    start = t.rfind('\r\n', 0, idx) + 2
    end_marker = '\r\nEnd Sub'
    end_idx = t.find(end_marker, idx)
    if end_idx == -1:
        return None
    end = end_idx + len(end_marker) + 2  # +2 para o \r\n após End Sub
    return t[start:end]

for sh in subs_to_remove:
    seg = extract_sub(text, sh)
    if seg is None:
        print(f'NAOACHEI sub: {sh}')
    else:
        c = text.count(seg)
        print(f'sub {sh}: count={c}')
        removals.append(seg)

# ── Remoção das 3 linhas de txtCodClientePedidos em cboClientePedidos_LostFocus ──
old_lostfocus_lines = (
    'If cboClientePedidos.Text = "" Then txtCodClientePedidos.Text = "": Exit Sub\r\n'
    'If cboClientePedidos.ListIndex = -1 Then txtCodClientePedidos.Text = "": Exit Sub\r\n'
    '\r\n'
    'txtCodClientePedidos = cboClientePedidos.ItemData(cboClientePedidos.ListIndex)\r\n'
    '\r\n'
    'TrataErro:\r\n'
    '   If Err.Number = 381 Then Exit Sub\r\n'
    'End Sub'
)
new_lostfocus_lines = (
    'TrataErro:\r\n'
    '   If Err.Number = 381 Then Exit Sub\r\n'
    'End Sub'
)
c_lf = text.count(old_lostfocus_lines)
print(f'lostfocus_lines: count={c_lf}')

# ── txtCodClientePedidos.Visible = False em cboIndicePedidos_Click ──────────
old_vis = 'txtCodClientePedidos.Visible = False\r\n'
c_vis = text.count(old_vis)
print(f'txtCodClientePedidos.Visible: count={c_vis}')

# ── Aplicar ──────────────────────────────────────────────────────────────────
all_ok = True
for seg in removals:
    if text.count(seg) != 1:
        print(f'ERRO count != 1 para trecho: {repr(seg[:60])}')
        all_ok = False

if c_lf != 1 or c_vis != 1:
    print(f'ERRO contagens: lostfocus={c_lf} vis={c_vis}')
    all_ok = False

if all_ok:
    for seg in removals:
        text = text.replace(seg, '')
    text = text.replace(old_lostfocus_lines, new_lostfocus_lines)
    text = text.replace(old_vis, '')

    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb').write(out)
    print('OK - todos os controles e subs removidos')
else:
    print('ABORTADO - verifique os erros acima')
