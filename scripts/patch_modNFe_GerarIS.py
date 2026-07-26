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

# Assinatura real da DLL (Object Browser):
#   GerarItensImpostoIS(CSTIS, ClassTribIS, vBCIS, PIS, pISESpec, uTrib, qTrib, vIS, alerta, erro)
#   1  CSTIS       = IS_CST
#   2  ClassTribIS = cClassTrib_IS
#   3  vBCIS       = IS_vBC
#   4  PIS         = IS_pAliq  (aliquota ad valorem %)
#   5  pISESpec    = IS_vUnid  (valor por unidade — Ad Rem)
#   6  uTrib       = UnidadeTributavel (string da unidade)
#   7  qTrib       = IS_qUnid (quantidade tributavel)
#   8  vIS         = IS_vIS
#
# Chamada errada apos patch anterior (params 5-7 fora de ordem):
#   ..., dISqUnid, CStr(dISvUnid), IIf(IsNull(NFeItens!IS_tipo_calculo)...), vIS, ...
#
# Chamada correta:
#   ..., dISvUnid, IIf(IsNull(NFeItens!UnidadeTributavel),"",NFeItens!UnidadeTributavel), dISqUnid, vIS, ...

old_call = (
    '              iRetorno = sistNFe.GerarItensImpostoIS('
    'IIf(IsNull(NFeItens!IS_CST), "000", NFeItens!IS_CST), '
    'IIf(IsNull(NFeItens!cClassTrib_IS), "000000", NFeItens!cClassTrib_IS), '
    'vBCIS, pIS, dISqUnid, CStr(dISvUnid), '
    'IIf(IsNull(NFeItens!IS_tipo_calculo), 0, CInt(NFeItens!IS_tipo_calculo)), '
    'vIS, mensagemAlerta, mensagemErro)'
)

new_call = (
    '              iRetorno = sistNFe.GerarItensImpostoIS('
    'IIf(IsNull(NFeItens!IS_CST), "99", NFeItens!IS_CST), '
    'IIf(IsNull(NFeItens!cClassTrib_IS), "", NFeItens!cClassTrib_IS), '
    'vBCIS, pIS, dISvUnid, '
    'IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), '
    'dISqUnid, '
    'vIS, mensagemAlerta, mensagemErro)'
)

content = sub('GerarItensImpostoIS params (2x)', old_call, new_call, content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
