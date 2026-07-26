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

# ── 1: SELECT totais — separar IBS_vBC e CBS_vBC ─────────────────────────────
content = sub('SELECT vBCCBSIBS split',
    '       "ISNULL(SUM(IBS_vBC),     0) AS vBCCBSIBS,    " & _' + R +
    '       "ISNULL(SUM(IBS_vIBSUF),  0) AS vIBSUF,        " & _',

    '       "ISNULL(SUM(IBS_vBC),     0) AS vBCIBS,        " & _' + R +
    '       "ISNULL(SUM(CBS_vBC),     0) AS vBCCBS,        " & _' + R +
    '       "ISNULL(SUM(IBS_vIBSUF),  0) AS vIBSUF,        " & _',
    content)

# ── 2: Dim — substituir varBCCBSIBS por varBCCBS e varBCIBS ──────────────────
content = sub('Dim varBCCBSIBS split',
    'Dim varBCCBSIBS As Double, varIBSUF As Double, varIBSMun As Double',
    'Dim varBCCBS As Double, varBCIBS As Double, varIBSUF As Double, varIBSMun As Double',
    content)

# ── 3: Leitura do recordset ───────────────────────────────────────────────────
content = sub('Read varBCCBSIBS split',
    '    varBCCBSIBS = ValidateNull(rTotais("vBCCBSIBS"))' + R +
    '    varIBSUF = ValidateNull(rTotais("vIBSUF"))',

    '    varBCCBS = ValidateNull(rTotais("vBCCBS"))' + R +
    '    varBCIBS = ValidateNull(rTotais("vBCIBS"))' + R +
    '    varIBSUF = ValidateNull(rTotais("vIBSUF"))',
    content)

# ── 4: UPDATE NotaFiscal ──────────────────────────────────────────────────────
content = sub('UPDATE vBCCBSIBS split',
    '       "vBCCBSIBS           = " & FSQL(varBCCBSIBS, 2) & ", " & _' + R +
    '       "vIBSUF              = " & FSQL(varIBSUF, 2) & ", " & _',

    '       "vBCCBS              = " & FSQL(varBCCBS, 2) & ", " & _' + R +
    '       "vBCIBS              = " & FSQL(varBCIBS, 2) & ", " & _' + R +
    '       "vIBSUF              = " & FSQL(varIBSUF, 2) & ", " & _',
    content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
