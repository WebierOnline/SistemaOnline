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

# GerarItensImpostoIBSCBS — assinatura completa (27 params + alerta + erro):
#   (1)  cst          = IBSCBS_CST
#   (2)  classTrib    = cClassTrib
#   (3)  vBC
# IBS-UF:
#   (4)  pAliqUF      = pIBSUF
#   (5)  pDifUF       = 0
#   (6)  vDifUF       = 0
#   (7)  vDevTribUF   = 0
#   (8)  pRedAliqUF   = pIBSpRed
#   (9)  pAliqEfetUF  = pIBSUF  (sem reducao: efetiva = nominal)
#   (10) vIBSUF
# IBS-Mun:
#   (11) pAliqMun     = pIBSMun
#   (12) pDifMun      = 0
#   (13) vDifMun      = 0
#   (14) vDevTribMun  = 0
#   (15) pRedAliqMun  = pIBSpRed
#   (16) pAliqEfetMun = pIBSMun  (sem reducao: efetiva = nominal)
#   (17) vIBSMun
#   (18) vIBS
# CBS:
#   (19) pAliqCBS     = pCBS
#   (20) pDifCBS      = 0
#   (21) vDifCBS      = 0
#   (22) vDevTribCBS  = 0
#   (23) pRedAliqCBS  = pCBSpRed
#   (24) pAliqEfetCBS = pCBS     (sem reducao: efetiva = nominal)
#   (25) vCBS

old_block = (
    '           If vBCCBSIBS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS('
    'IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), '
    'IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), '
    'vBCCBSIBS, pIBSUF, 0, 0, 0, pIBSpRed, 0, vIBSUF, '
    'pIBSMun, 0, 0, 0, pIBSpRed, 0, vIBSMun, vIBS, '
    'pCBS, 0, 0, 0, pCBSpRed, 0, vCBS, '
    'mensagemAlerta, mensagemErro)' + R +
    '           End If'
)

new_block = (
    '           If vBCCBSIBS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS( _' + R +
    '                  IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), _' + R +
    '                  IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), _' + R +
    '                  vBCCBSIBS, _' + R +
    '                  pIBSUF, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  pIBSpRed, _' + R +
    '                  pIBSUF, _' + R +
    '                  vIBSUF, _' + R +
    '                  pIBSMun, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  pIBSpRed, _' + R +
    '                  pIBSMun, _' + R +
    '                  vIBSMun, _' + R +
    '                  vIBS, _' + R +
    '                  pCBS, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  0, _' + R +
    '                  pCBSpRed, _' + R +
    '                  pCBS, _' + R +
    '                  vCBS, _' + R +
    '                  mensagemAlerta, _' + R +
    '                  mensagemErro)' + R +
    '           End If'
)

content = sub('GerarItensImpostoIBSCBS layout+pAliqEfet (2x)', old_block, new_block, content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
