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

# VB6 permite no maximo 25 continuacoes (_) por linha logica.
# O bloco anterior tinha 27 — excede o limite e impede o carregamento do modulo.
# Solucao: agrupar por bloco (IBS-UF / IBS-Mun / CBS), usando 8 continuacoes.
#
# Assinatura GerarItensImpostoIBSCBS (27 params + alerta + erro):
#   (1) cst  (2) classTrib  (3) vBC
#   IBS-UF:  (4)pAliqUF (5)pDif=0 (6)vDif=0 (7)vDevTrib=0 (8)pRedAliq=pIBSpRed (9)pAliqEfet=pIBSUF (10)vIBSUF
#   IBS-Mun: (11)pAliqMun (12)pDif=0 (13)vDif=0 (14)vDevTrib=0 (15)pRedAliq=pIBSpRed (16)pAliqEfet=pIBSMun (17)vIBSMun (18)vIBS
#   CBS:     (19)pAliqCBS (20)pDif=0 (21)vDif=0 (22)vDevTrib=0 (23)pRedAliq=pCBSpRed (24)pAliqEfet=pCBS (25)vCBS

# Substituir o bloco com 27 _ pelo bloco agrupado com 8 _
old_block = (
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

new_block = (
    '           If vBCCBSIBS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS( _' + R +
    '                  IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), _' + R +
    '                  IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), _' + R +
    '                  vBCCBSIBS, _' + R +
    '                  pIBSUF, 0, 0, 0, pIBSpRed, pIBSUF, vIBSUF, _' + R +
    '                  pIBSMun, 0, 0, 0, pIBSpRed, pIBSMun, vIBSMun, vIBS, _' + R +
    '                  pCBS, 0, 0, 0, pCBSpRed, pCBS, vCBS, _' + R +
    '                  mensagemAlerta, mensagemErro)' + R +
    '           End If'
)

content = sub('GerarItensImpostoIBSCBS layout agrupado (2x)', old_block, new_block, content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
