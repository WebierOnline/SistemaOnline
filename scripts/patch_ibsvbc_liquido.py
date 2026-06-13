# -*- coding: utf-8 -*-
# Corrige RecalcularItensNota: IBS_vBC usa valor liquido (vProd - vDesconto) em vez do bruto.
# Isso corrige simultaneamente:
#   - vBC no item XML (de 4,00 para 3,50 com desconto de 0,50)
#   - vNFTot nos totais (TotvBCCBSIBS passa a ser liquido, corrigindo 4,01 -> 3,51)
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: adicionar ValorDesconto ao SELECT ────────────────────────────────
patches.append(('select_valordesconto',
    '"IBS_UFpAliq, IBS_MunpAliq, CBS_pAliq, IS_pAliq, IBS_pRed, CBS_pRed " & _\r\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota ',
    '"IBS_UFpAliq, IBS_MunpAliq, CBS_pAliq, IS_pAliq, IBS_pRed, CBS_pRed, ValorDesconto " & _\r\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota ',
    1))

# ── Patch 2: adicionar Dim curDesconto2 As Currency nas declaracoes ───────────
patches.append(('dim_curDesconto2',
    'Dim curBCCBSIBS2 As Currency\r\nDim curVIBSUF2   As Currency, curVIBSMun2 As Currency, curVCBS2 As Currency\r\n',
    'Dim curBCCBSIBS2 As Currency\r\nDim curDesconto2 As Currency\r\nDim curVIBSUF2   As Currency, curVIBSMun2 As Currency, curVCBS2 As Currency\r\n',
    1))

# ── Patch 3: ler ValorDesconto e usar valor liquido na BC do IBS/CBS ──────────
patches.append(('bccbsibs_liquido',
    '    curBCCBSIBS2 = vValProd\r\n',
    '    curDesconto2 = CCur(IIf(IsNull(rItens("ValorDesconto")), 0, rItens("ValorDesconto")))\r\n'
    '    curBCCBSIBS2 = vValProd - curDesconto2\r\n'
    '    If curBCCBSIBS2 < 0 Then curBCCBSIBS2 = 0\r\n',
    1))

all_ok = True
for name, old, new, expected in patches:
    c = text.count(old)
    ok = (c == expected)
    print(f'{name}: {"OK" if ok else "ERRO"} count={c} esperado={expected}')
    if not ok:
        all_ok = False

if all_ok:
    for _, old, new, _ in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK — IBS_vBC agora usa valor liquido (vProd - vDesconto)')
else:
    print('\nABORTADO')
