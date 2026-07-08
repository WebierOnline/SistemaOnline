path = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c, expected=1):
    n = c.count(old)
    if n != expected:
        errors.append(f'{label}: count={n} (esperado {expected})')
        return c
    print(f'{label} OK (n={n})')
    return c.replace(old, new)

# ── 1: Adicionar pIBSpRed e pCBSpRed nas Dims da funcao ──────────────────────
content = sub('Dim pIBSpRed/pCBSpRed',
    'Dim dISqUnid As Double, dISvUnid As Double',
    'Dim dISqUnid As Double, dISvUnid As Double' + R +
    'Dim pIBSpRed As Double, pCBSpRed As Double',
    content)

# ── 2: Ler pIBSpRed e pCBSpRed do recordset (aparece 2x — ambas as branches) ─
content = sub('Read pIBSpRed/pCBSpRed (2x)',
    '           vCBS = IIf(IsNull(NFeItens!CBS_vCBS), 0, CDbl(NFeItens!CBS_vCBS))' + R +
    '           If vBCCBSIBS > 0 Then',

    '           vCBS = IIf(IsNull(NFeItens!CBS_vCBS), 0, CDbl(NFeItens!CBS_vCBS))' + R +
    '           pIBSpRed = IIf(IsNull(NFeItens!IBS_pRed), 0, CDbl(NFeItens!IBS_pRed))' + R +
    '           pCBSpRed = IIf(IsNull(NFeItens!CBS_pRed), 0, CDbl(NFeItens!CBS_pRed))' + R +
    '           If vBCCBSIBS > 0 Then',
    content, expected=2)

# ── 3: Passar pIBSpRed/pCBSpRed nas posicoes corretas da chamada (2x) ─────────
# Assinatura completa:
#   cst, classTrib, vBC,
#   pIBSUF(4), pDifUF(5)=0, vDifUF(6)=0, vDevTribUF(7)=0, pRedAliqUF(8)=pIBSpRed, pAliqEfetUF(9)=0, vIBSUF(10),
#   pIBSMun(11), pDifMun(12)=0, vDifMun(13)=0, vDevTribMun(14)=0, pRedAliqMun(15)=pIBSpRed, pAliqEfetMun(16)=0, vIBSMun(17), vIBS(18),
#   pCBS(19), pDifCBS(20)=0, vDifCBS(21)=0, vDevTribCBS(22)=0, pRedAliqCBS(23)=pCBSpRed, pAliqEfetCBS(24)=0, vCBS(25)
old_call = (
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS('
    'IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), '
    'IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), '
    'vBCCBSIBS, pIBSUF, 0, 0, 0, 0, 0, vIBSUF, '
    'pIBSMun, 0, 0, 0, 0, 0, vIBSMun, vIBS, '
    'pCBS, 0, 0, 0, 0, 0, vCBS, '
    'mensagemAlerta, mensagemErro)'
)

new_call = (
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS('
    'IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), '
    'IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), '
    'vBCCBSIBS, pIBSUF, 0, 0, 0, pIBSpRed, 0, vIBSUF, '
    'pIBSMun, 0, 0, 0, pIBSpRed, 0, vIBSMun, vIBS, '
    'pCBS, 0, 0, 0, pCBSpRed, 0, vCBS, '
    'mensagemAlerta, mensagemErro)'
)

content = sub('GerarItensImpostoIBSCBS pRed (2x)', old_call, new_call, content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
