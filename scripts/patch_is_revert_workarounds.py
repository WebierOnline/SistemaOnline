# -*- coding: utf-8 -*-
# Reverte os 3 workarounds aplicados para contornar a limitacao da DLL antiga.
# A DLL nova corrige ShouldSerializeqTrib(), entao IS tipo=1 (ad valorem) agora
# gera <vBCIS>+<pIS>+<vIS> sem <qTrib>0.0000</qTrib> invalido.
#
# Reversoes:
#  1. modNFe.bas: "If vIS > 0 And dISqUnid > 0" -> "If vIS > 0"  (2 ocorrencias)
#  2. modNFe.bas: "If dISqUnid > 0 Then TotvIS = ..." -> "TotvIS = ..."
#  3. NFe_Completa.frm AtualizarTotaisNota: SUM(IS_vIS) sem filtro IS_qUnid

import sys

# ══════════════════════════════════════════════════════════════════
# Patch 1 e 2: modNFe.bas
# ══════════════════════════════════════════════════════════════════
PATH1 = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
text1 = open(PATH1, 'rb').read().decode('windows-1252')
ok1 = True

OLD_COND = '           If vIS > 0 And dISqUnid > 0 Then\r\n'
NEW_COND = '           If vIS > 0 Then\r\n'
c1 = text1.count(OLD_COND)
print(f'Patch 1 (condicao GerarItensImpostoIS): {c1} ocorrencias (esperado: 2)')
if c1 == 2:
    text1 = text1.replace(OLD_COND, NEW_COND)
else:
    print('  ABORTADO'); ok1 = False

OLD_TOT = '        If dISqUnid > 0 Then TotvIS = TotvIS + vIS\r\n'
NEW_TOT = '        TotvIS = TotvIS + vIS\r\n'
c2 = text1.count(OLD_TOT)
print(f'Patch 2 (TotvIS acumulador): {c2} ocorrencias (esperado: 1)')
if c2 == 1:
    text1 = text1.replace(OLD_TOT, NEW_TOT)
else:
    print('  ABORTADO'); ok1 = False

if ok1:
    out1 = text1.encode('windows-1252')
    out1 = out1.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH1, 'wb').write(out1)
    print('modNFe.bas: OK')
else:
    sys.exit(1)

# ══════════════════════════════════════════════════════════════════
# Patch 3: NFe_Completa.frm — AtualizarTotaisNota
# ══════════════════════════════════════════════════════════════════
PATH2 = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
text2 = open(PATH2, 'rb').read().decode('windows-1252')

OLD_SQL = (
    '"ISNULL(SUM(IS_vBC),      0) AS vBCIS,         " & _\r\n'
    '       "ISNULL(SUM(CASE WHEN ISNULL(IS_qUnid,0) > 0 THEN ISNULL(IS_vIS,0) ELSE 0 END), 0) AS vIS " & _\r\n'
)
NEW_SQL = (
    '"ISNULL(SUM(IS_vBC),      0) AS vBCIS,         " & _\r\n'
    '       "ISNULL(SUM(IS_vIS),      0) AS vIS            " & _\r\n'
)
c3 = text2.count(OLD_SQL)
print(f'\nPatch 3 (AtualizarTotaisNota SUM IS_vIS): {c3} ocorrencias (esperado: 1)')
if c3 == 1:
    text2 = text2.replace(OLD_SQL, NEW_SQL)
    out2 = text2.encode('windows-1252')
    out2 = out2.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH2, 'wb').write(out2)
    print('NFe_Completa.frm: OK')
else:
    print('NFe_Completa.frm: ABORTADO'); sys.exit(1)

print('\n=== Revertido ===')
print('IS tipo=1 (ad valorem): agora gera <vBCIS>+<pIS>+<vIS> sem <qTrib> invalido')
print('IS tipo=2 (ad rem): gera <pISEspec>+<uTrib>+<qTrib>+<vIS> como antes')
print('IS tipo=3 (misto): gera todos os elementos')
print('ValorNota e ISTot.vIS incluem IS de TODOS os tipos')
