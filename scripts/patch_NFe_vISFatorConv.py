path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Declarar vISFatorConv no modulo ───────────────────────────────────────
content = sub('Dim vISFatorConv',
    'Dim vISvUnid As String          \'bloco IS',
    'Dim vISvUnid As String          \'bloco IS' + R +
    'Dim vISFatorConv As String      \'bloco IS',
    content)

# ── 2: Acrescentar fator_conversao_IS no SELECT de produtos ──────────────────
content = sub('SELECT fator_conversao_IS',
    'IBSCBSCST, ISCST, cClassTrib_IS,',
    'IBSCBSCST, ISCST, cClassTrib_IS, fator_conversao_IS,',
    content)

# ── 3: Ler r("fator_conversao_IS") apos o bloco IS ───────────────────────────
content = sub('read vISFatorConv',
    '        Set rIS = Nothing' + R +
    '     End If' + R +
    '     ' + R +
    R +
    '     ',
    '        Set rIS = Nothing' + R +
    '     End If' + R +
    '     vISFatorConv = Format(ValidateNull(r("fator_conversao_IS")), "##,##0.000")' + R +
    '     ' + R +
    R +
    '     ',
    content)

# ── 4a: Limpar no bloco Else (produto nao encontrado) ────────────────────────
content = sub('Else clear vISFatorConv',
    '     vISCST      = ""' + R +
    '     vISTipoCalc = ""' + R +
    '     vISpAliq    = ""' + R +
    '     vISqUnid    = ""' + R +
    '     vISvUnid    = ""' + R +
    ' End If',
    '     vISCST       = ""' + R +
    '     vISTipoCalc  = ""' + R +
    '     vISpAliq     = ""' + R +
    '     vISqUnid     = ""' + R +
    '     vISvUnid     = ""' + R +
    '     vISFatorConv = ""' + R +
    ' End If',
    content)

# ── 4b: Limpar em LimparObjetosProduto ───────────────────────────────────────
content = sub('LimparObjetos clear vISFatorConv',
    'vISCST      = ""' + R +
    'vISTipoCalc = ""' + R +
    'vISpAliq    = ""' + R +
    'vISqUnid    = ""' + R +
    'vISvUnid    = ""' + R +
    'End Sub',
    'vISCST       = ""' + R +
    'vISTipoCalc  = ""' + R +
    'vISpAliq     = ""' + R +
    'vISqUnid     = ""' + R +
    'vISvUnid     = ""' + R +
    'vISFatorConv = ""' + R +
    'End Sub',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
