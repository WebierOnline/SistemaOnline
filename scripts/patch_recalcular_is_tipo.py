# -*- coding: utf-8 -*-
# Corrige RecalcularItensNota em NFe_Completa.frm:
# Adiciona IS_tipo_calculo, IS_qUnid, IS_vUnid ao SELECT e calcula IS_vIS
# corretamente por tipo (0=zero, 1=ad valorem, 2=ad rem, 3=mista).
# Antes: usava apenas IS_pAliq, zerando IS_vIS de itens ad rem (tipo 2).
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: adicionar IS_tipo_calculo, IS_qUnid, IS_vUnid ao SELECT ─────────
patches.append(('select_is_tipo',
    '"IBS_UFpAliq, IBS_MunpAliq, CBS_pAliq, IS_pAliq, IBS_pRed, CBS_pRed, ValorDesconto " & _\r\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota ',
    '"IBS_UFpAliq, IBS_MunpAliq, CBS_pAliq, IS_pAliq, IBS_pRed, CBS_pRed, ValorDesconto, " & _\r\n'
    '       "IS_tipo_calculo, IS_qUnid, IS_vUnid " & _\r\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota ',
    1))

# ── Patch 2: adicionar Dim para novas variaveis IS ───────────────────────────
patches.append(('dim_is_tipo',
    'Dim curBCIS2     As Currency, curVIS2 As Currency\r\n',
    'Dim curBCIS2     As Currency, curVIS2 As Currency\r\n'
    'Dim iTipoIS2     As Integer\r\n'
    'Dim curISqUnid2  As Currency, curISvUnid2 As Currency\r\n',
    1))

# ── Patch 3: substituir calculo IS simples por Select Case por tipo ───────────
patches.append(('is_calc_tipo',
    '    curBCIS2 = vValProd\r\n'
    '    curVIS2 = CCur(Format(curBCIS2 * dblPIS / 100, "0.00"))\r\n',
    '    iTipoIS2 = IIf(IsNull(rItens("IS_tipo_calculo")), 0, CInt(rItens("IS_tipo_calculo")))\r\n'
    '    curISqUnid2 = CCur(IIf(IsNull(rItens("IS_qUnid")), 0, rItens("IS_qUnid")))\r\n'
    '    curISvUnid2 = CCur(IIf(IsNull(rItens("IS_vUnid")), 0, rItens("IS_vUnid")))\r\n'
    '    Select Case iTipoIS2\r\n'
    '        Case 1\r\n'
    '            curBCIS2 = vValProd\r\n'
    '            curVIS2 = CCur(Format(curBCIS2 * dblPIS / 100, "0.00"))\r\n'
    '        Case 2\r\n'
    '            curBCIS2 = 0\r\n'
    '            curVIS2 = CCur(Format(curISqUnid2 * curISvUnid2, "0.00"))\r\n'
    '        Case 3\r\n'
    '            curBCIS2 = vValProd\r\n'
    '            curVIS2 = CCur(Format(curBCIS2 * dblPIS / 100, "0.00")) + CCur(Format(curISqUnid2 * curISvUnid2, "0.00"))\r\n'
    '        Case Else\r\n'
    '            curBCIS2 = 0\r\n'
    '            curVIS2 = 0\r\n'
    '    End Select\r\n',
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
    print('\nOK — RecalcularItensNota agora calcula IS por tipo (0/1/2/3)')
else:
    print('\nABORTADO')
