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

# ── 1: Adicionar dISqUnid, dISvUnid nas Dims da funcao ───────────────────────
content = sub('Dim dISqUnid/dISvUnid',
    'Dim vBCIS As Double, pIS As Double, vIS As Double',
    'Dim vBCIS As Double, pIS As Double, vIS As Double' + R +
    'Dim dISqUnid As Double, dISvUnid As Double',
    content)

# ── 2: Bloco IBS/CBS/IS (identico nas 2 branches: nao-combustivel e combustivel)
# Substituir referencias de coluna erradas + condicao IS + qUnid/vUnid
old_block = (
    "                      ' IBS/CBS: todos os regimes" + R +
    '           vBCCBSIBS = IIf(IsNull(NFeItens!vBCCBSIBS), 0, CDbl(NFeItens!vBCCBSIBS))' + R +
    '           pIBSUF = IIf(IsNull(NFeItens!IBSUFpAliq), 0, CDbl(NFeItens!IBSUFpAliq))' + R +
    '           vIBSUF = IIf(IsNull(NFeItens!vIBSUF), 0, CDbl(NFeItens!vIBSUF))' + R +
    '           pIBSMun = IIf(IsNull(NFeItens!IBSMunpAliq), 0, CDbl(NFeItens!IBSMunpAliq))' + R +
    '           vIBSMun = IIf(IsNull(NFeItens!vIBSMun), 0, CDbl(NFeItens!vIBSMun))' + R +
    '           vIBS = vIBSUF + vIBSMun' + R +
    '           pCBS = IIf(IsNull(NFeItens!CBSpAliq), 0, CDbl(NFeItens!CBSpAliq))' + R +
    '           vCBS = IIf(IsNull(NFeItens!vCBS), 0, CDbl(NFeItens!vCBS))' + R +
    '           If vBCCBSIBS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS(IIf(IsNull(NFeItens!IBSCBSCST), "000", NFeItens!IBSCBSCST), "000001", vBCCBSIBS, pIBSUF, 0, 0, 0, 0, 0, vIBSUF, pIBSMun, 0, 0, 0, 0, 0, vIBSMun, vIBS, pCBS, 0, 0, 0, 0, 0, vCBS, mensagemAlerta, mensagemErro)' + R +
    '           End If' + R +
    "           ' IS (Imposto Seletivo)" + R +
    '           vBCIS = IIf(IsNull(NFeItens!vBCIS), 0, CDbl(NFeItens!vBCIS))' + R +
    '           pIS = IIf(IsNull(NFeItens!ISpAliq), 0, CDbl(NFeItens!ISpAliq))' + R +
    '           vIS = IIf(IsNull(NFeItens!vIS), 0, CDbl(NFeItens!vIS))' + R +
    '           If vBCIS > 0 And vIS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIS(IIf(IsNull(NFeItens!ISCST), "000", NFeItens!ISCST), 0, vBCIS, pIS, 0, "", 0, vIS, mensagemAlerta, mensagemErro)' + R +
    '           End If'
)

new_block = (
    "                      ' IBS/CBS: todos os regimes" + R +
    '           vBCCBSIBS = IIf(IsNull(NFeItens!IBS_vBC), 0, CDbl(NFeItens!IBS_vBC))' + R +
    '           pIBSUF = IIf(IsNull(NFeItens!IBS_UFpAliq), 0, CDbl(NFeItens!IBS_UFpAliq))' + R +
    '           vIBSUF = IIf(IsNull(NFeItens!IBS_vIBSUF), 0, CDbl(NFeItens!IBS_vIBSUF))' + R +
    '           pIBSMun = IIf(IsNull(NFeItens!IBS_MunpAliq), 0, CDbl(NFeItens!IBS_MunpAliq))' + R +
    '           vIBSMun = IIf(IsNull(NFeItens!IBS_vIBSMun), 0, CDbl(NFeItens!IBS_vIBSMun))' + R +
    '           vIBS = vIBSUF + vIBSMun' + R +
    '           pCBS = IIf(IsNull(NFeItens!CBS_pAliq), 0, CDbl(NFeItens!CBS_pAliq))' + R +
    '           vCBS = IIf(IsNull(NFeItens!CBS_vCBS), 0, CDbl(NFeItens!CBS_vCBS))' + R +
    '           If vBCCBSIBS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIBSCBS(IIf(IsNull(NFeItens!IBSCBS_CST), "000", NFeItens!IBSCBS_CST), IIf(IsNull(NFeItens!cClassTrib), "000001", NFeItens!cClassTrib), vBCCBSIBS, pIBSUF, 0, 0, 0, 0, 0, vIBSUF, pIBSMun, 0, 0, 0, 0, 0, vIBSMun, vIBS, pCBS, 0, 0, 0, 0, 0, vCBS, mensagemAlerta, mensagemErro)' + R +
    '           End If' + R +
    "           ' IS (Imposto Seletivo)" + R +
    '           vBCIS = IIf(IsNull(NFeItens!IS_vBC), 0, CDbl(NFeItens!IS_vBC))' + R +
    '           pIS = IIf(IsNull(NFeItens!IS_pAliq), 0, CDbl(NFeItens!IS_pAliq))' + R +
    '           vIS = IIf(IsNull(NFeItens!IS_vIS), 0, CDbl(NFeItens!IS_vIS))' + R +
    '           dISqUnid = IIf(IsNull(NFeItens!IS_qUnid), 0, CDbl(NFeItens!IS_qUnid))' + R +
    '           dISvUnid = IIf(IsNull(NFeItens!IS_vUnid), 0, CDbl(NFeItens!IS_vUnid))' + R +
    '           If vIS > 0 Then' + R +
    '              iRetorno = sistNFe.GerarItensImpostoIS(IIf(IsNull(NFeItens!IS_CST), "000", NFeItens!IS_CST), IIf(IsNull(NFeItens!cClassTrib_IS), "000000", NFeItens!cClassTrib_IS), vBCIS, pIS, dISqUnid, CStr(dISvUnid), IIf(IsNull(NFeItens!IS_tipo_calculo), 0, CInt(NFeItens!IS_tipo_calculo)), vIS, mensagemAlerta, mensagemErro)' + R +
    '           End If'
)

content = sub('bloco IBS/CBS/IS (2x)', old_block, new_block, content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
