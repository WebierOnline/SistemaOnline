path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []

def sub(label, old, new, c, replace_all=False):
    n = c.count(old)
    if replace_all:
        if n == 0:
            errors.append(f'{label}: count=0')
            return c
        print(f'{label} OK ({n}x)')
        return c.replace(old, new)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: SELECT NotaFiscalItens — nomes das colunas lidas no RecalcularItensNota ─
content = sub('SELECT NFI cols',
    '"IBSUFpAliq, IBSMunpAliq, CBSpAliq, ISpAliq "',
    '"IBS_UFpAliq, IBS_MunpAliq, CBS_pAliq, IS_pAliq "',
    content)

# ── 2: rItens() — acesso ao Recordset de NotaFiscalItens ─────────────────────
# (preserve r("IBSUFpAliq") da tabela produtos — variavel diferente)
content = sub('rItens IBS_UFpAliq',   'rItens("IBSUFpAliq")',  'rItens("IBS_UFpAliq")',  content, replace_all=True)
content = sub('rItens IBS_MunpAliq',  'rItens("IBSMunpAliq")', 'rItens("IBS_MunpAliq")', content, replace_all=True)
content = sub('rItens CBS_pAliq',     'rItens("CBSpAliq")',    'rItens("CBS_pAliq")',    content, replace_all=True)
content = sub('rItens IS_pAliq',      'rItens("ISpAliq")',     'rItens("IS_pAliq")',     content, replace_all=True)

# ── 3: UPDATE NotaFiscalItens SET — nomes sem espacos (distintos do UPDATE NotaFiscal) ──
content = sub('UPD NFI IBS_vBC',    '"vBCCBSIBS = "', '"IBS_vBC = "',    content)
content = sub('UPD NFI IBS_vIBSUF', '"vIBSUF = "',    '"IBS_vIBSUF = "', content)
content = sub('UPD NFI IBS_vIBSMun','"vIBSMun = "',   '"IBS_vIBSMun = "',content)
content = sub('UPD NFI IBS_vIBS',   '"vIBS = "',      '"IBS_vIBS = "',   content)
content = sub('UPD NFI CBS_vCBS',   '"vCBS = "',      '"CBS_vCBS = "',   content)
content = sub('UPD NFI IS_vBC',     '"vBCIS = "',     '"IS_vBC = "',     content)
content = sub('UPD NFI IS_vIS',     '"vIS = "',       '"IS_vIS = "',     content)

# ── 4: SUM query — coluna real atualizada; alias preservado (rTotais / UPDATE NotaFiscal inalterados) ──
content = sub('SUM IBS_vBC',    'ISNULL(SUM(vBCCBSIBS),  0) AS vBCCBSIBS,    ', 'ISNULL(SUM(IBS_vBC),     0) AS vBCCBSIBS,    ', content)
content = sub('SUM IBS_vIBSUF', 'ISNULL(SUM(vIBSUF),     0) AS vIBSUF,        ', 'ISNULL(SUM(IBS_vIBSUF),  0) AS vIBSUF,        ', content)
content = sub('SUM IBS_vIBSMun','ISNULL(SUM(vIBSMun),    0) AS vIBSMun,       ', 'ISNULL(SUM(IBS_vIBSMun), 0) AS vIBSMun,       ', content)
content = sub('SUM IBS_vIBS',   'ISNULL(SUM(vIBS),       0) AS vIBS,          ', 'ISNULL(SUM(IBS_vIBS),    0) AS vIBS,          ', content)
content = sub('SUM CBS_vCBS',   'ISNULL(SUM(vCBS),       0) AS vCBS,          ', 'ISNULL(SUM(CBS_vCBS),    0) AS vCBS,          ', content)
content = sub('SUM IS_vBC',     'ISNULL(SUM(vBCIS),      0) AS vBCIS,         ', 'ISNULL(SUM(IS_vBC),      0) AS vBCIS,         ', content)
content = sub('SUM IS_vIS',     'ISNULL(SUM(vIS),        0) AS vIS            ', 'ISNULL(SUM(IS_vIS),      0) AS vIS            ', content)

# ── 5: Tb() — INSERT NotaFiscalItens ─────────────────────────────────────────
content = sub('Tb IBSCBS_CST',   'Tb("IBSCBSCST")',  'Tb("IBSCBS_CST")',  content)
content = sub('Tb CBS_pAliq',    'Tb("CBSpAliq")',   'Tb("CBS_pAliq")',   content)
content = sub('Tb IBS_UFpAliq',  'Tb("IBSUFpAliq")', 'Tb("IBS_UFpAliq")', content)
content = sub('Tb IBS_MunpAliq', 'Tb("IBSMunpAliq")','Tb("IBS_MunpAliq")',content)
content = sub('Tb IBS_vBC',      'Tb("vBCCBSIBS")',  'Tb("IBS_vBC")',     content)
content = sub('Tb IBS_vIBSUF',   'Tb("vIBSUF")',     'Tb("IBS_vIBSUF")',  content)
content = sub('Tb IBS_vIBSMun',  'Tb("vIBSMun")',    'Tb("IBS_vIBSMun")', content)
content = sub('Tb IBS_vIBS',     'Tb("vIBS")',       'Tb("IBS_vIBS")',    content)
content = sub('Tb CBS_vCBS',     'Tb("vCBS")',       'Tb("CBS_vCBS")',    content)
content = sub('Tb IS_CST',       'Tb("ISCST")',      'Tb("IS_CST")',      content)
content = sub('Tb IS_pAliq',     'Tb("ISpAliq")',    'Tb("IS_pAliq")',    content)
content = sub('Tb IS_vBC',       'Tb("vBCIS")',      'Tb("IS_vBC")',      content)
content = sub('Tb IS_vIS',       'Tb("vIS")',        'Tb("IS_vIS")',      content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
