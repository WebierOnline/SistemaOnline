# -*- coding: utf-8 -*-
# Remove o botão cmdEnviarPDF do form NFe_Completa:
# 1. Lista Tab(1).Control do SSTab (renumera controles)
# 2. Bloco Begin...End da definição do controle
# 3. Private Sub cmdEnviarPDF_Click() ... End Sub
# 4. Todas as linhas cmdEnviarPDF.Enabled = True/False
import re

PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []  # (name, old, new, expected_count)

# ── Patch 1: SSTab Tab(1).Control — remover cmdEnviarPDF e renumerar 1-17→0-16
INDENT = '      '
controls = [
    'cmdEnviarPDF',   # 0 — será removido
    'cmdEnviarXML',   # 1 → 0
    'cmdEspelho',     # 2 → 1
    'cmdEditar',      # 3 → 2
    'cmdCartaCorrecao',# 4 → 3
    'cmdInutilizar',  # 5 → 4
    'cmdDuplicar',    # 6 → 5
    'cmdConsultar',   # 7 → 6
    'cmdCancelarNota',# 8 → 7
    'cmdTransmitir',  # 9 → 8
    'cmdImprimir',    # 10 → 9
    'cmdCopiarChave', # 11 → 10
    'GridNotas',      # 12 → 11
    'Frame4',         # 13 → 12
    'lvwTotais',      # 14 → 13
    'picAguarde',     # 15 → 14
    'frmCorre\xe7\xe3o', # 16 → 15  (frmCorreção)
    'txtCodObservacao',  # 17 → 16
]

def make_ctrl_block(ctrl_list, start_idx=0, count=None):
    lines = []
    for n, name in enumerate(ctrl_list, start=start_idx):
        lines.append(f"{INDENT}Tab(1).Control({n})=   \"{name}\"\r\n")
        lines.append(f"{INDENT}Tab(1).Control({n}).Enabled=   0   'False\r\n")
    cnt = count if count is not None else len(ctrl_list)
    lines.append(f"{INDENT}Tab(1).ControlCount=   {cnt}\r\n")
    return ''.join(lines)

old_ctrl = make_ctrl_block(controls, start_idx=0, count=18)
new_ctrl = make_ctrl_block([c for c in controls if c != 'cmdEnviarPDF'], start_idx=0, count=17)
patches.append(('sstab_controles', old_ctrl, new_ctrl, 1))

# ── Patch 2: Bloco Begin ChamaleonBtn.chameleonButton cmdEnviarPDF ... End ───
begin_marker = '\r\n      Begin ChamaleonBtn.chameleonButton cmdEnviarPDF '
i_begin = text.find(begin_marker)
assert i_begin > 0, 'Begin cmdEnviarPDF não encontrado'
i_end = text.find('\r\n      End\r\n', i_begin) + len('\r\n      End\r\n')
old_begin = text[i_begin:i_end]
patches.append(('begin_end_ctrl', old_begin, '', 1))

# ── Patch 3: Private Sub cmdEnviarPDF_Click() ... End Sub ─────────────────────
sub_marker = '\r\n\r\n\r\nPrivate Sub cmdEnviarPDF_Click()\r\n'
i_sub = text.find(sub_marker)
assert i_sub > 0, 'Sub cmdEnviarPDF_Click não encontrado'
i_sub_end = text.find('\r\nEnd Sub\r\n', i_sub) + len('\r\nEnd Sub\r\n')
old_sub = text[i_sub:i_sub_end]
patches.append(('sub_click', old_sub, '\r\n', 1))

# ── Patch 4: Linhas cmdEnviarPDF.Enabled = ... (todos os contextos) ───────────
patches.append(('enabled_false_noindent',
    'cmdEnviarXML.Enabled = False\r\ncmdEnviarPDF.Enabled = False\r\n',
    'cmdEnviarXML.Enabled = False\r\n', 1))
patches.append(('enabled_false_indent',   '    cmdEnviarPDF.Enabled = False\r\n', '', 3))
patches.append(('enabled_true_indent',    '    cmdEnviarPDF.Enabled = True\r\n',  '', 2))

# ── Verificar e aplicar ───────────────────────────────────────────────────────
all_ok = True
for name, old, new, expected in patches:
    c = text.count(old)
    ok = (c == expected)
    status = f'OK (count={c})' if ok else f'ERRO count={c} esperado={expected}'
    print(f'{name}: {status}')
    if not ok:
        all_ok = False

if all_ok:
    for _, old, new, _ in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK — cmdEnviarPDF removido do form')
else:
    print('\nABORTADO')
