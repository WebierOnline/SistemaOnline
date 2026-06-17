# -*- coding: utf-8 -*-
# Correcao final IS tipo=1 (ad valorem) — dois arquivos:
#
# PROBLEMA: itens IS tipo=1 (cervejas, IS_qUnid=0) causam:
#   - schema error: DLL sempre gera <qTrib>0.0000</qTrib> quando dISqUnid=0
#   - SEFAZ error: ISTot.vIS=70 mas per-item IS sum=66,40 se nao acumula em TotvIS
#
# SOLUCAO: Excluir IS tipo=1 de TODOS os calculos de total:
#   1. modNFe.bas: restaurar "And dISqUnid > 0" na condicao GerarItensImpostoIS
#   2. modNFe.bas: so acumular TotvIS quando dISqUnid > 0
#   3. NFe_Completa.frm AtualizarTotaisNota: so somar IS_vIS onde IS_qUnid > 0
#
# Resultado: ISTot.vIS=66,40 = per-item IS sum, ICMSTot.vNF=237,79, vNFTot=237,79
# IS tipo=2/3 (cigarros, refrig ad rem) continuam normais.

import sys

# ══════════════════════════════════════════════════════════════════
# Patch 1: modNFe.bas
# ══════════════════════════════════════════════════════════════════
PATH1 = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data1 = open(PATH1, 'rb').read()
text1 = data1.decode('windows-1252')
ok1 = True

# 1a) Restaurar "And dISqUnid > 0" nas duas ocorrencias (revert do patch_is_cond_revert)
OLD_COND = '           If vIS > 0 Then\r\n'
NEW_COND = '           If vIS > 0 And dISqUnid > 0 Then\r\n'
c = text1.count(OLD_COND)
print(f'modNFe patch 1a (condicao): {c} ocorrencias (esperado: 2)')
if c == 2:
    text1 = text1.replace(OLD_COND, NEW_COND)
else:
    print('  AVISO: nao aplicado (possivelmente ja estava correto)')
    ok1 = False

# 1b) TotvIS so acumula quando dISqUnid > 0
OLD_TOT = '        TotvIS = TotvIS + vIS\r\n'
NEW_TOT = '        If dISqUnid > 0 Then TotvIS = TotvIS + vIS\r\n'
c2 = text1.count(OLD_TOT)
print(f'modNFe patch 1b (TotvIS): {c2} ocorrencias (esperado: 1)')
if c2 == 1:
    text1 = text1.replace(OLD_TOT, NEW_TOT)
else:
    print('  AVISO: nao encontrado / ja modificado')
    ok1 = False

if ok1 or c == 2 or c2 == 1:
    out1 = text1.encode('windows-1252')
    out1 = out1.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH1, 'wb').write(out1)
    print('modNFe.bas: OK')

# ══════════════════════════════════════════════════════════════════
# Patch 2: NFe_Completa.frm — AtualizarTotaisNota
# ══════════════════════════════════════════════════════════════════
PATH2 = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data2 = open(PATH2, 'rb').read()
text2 = data2.decode('windows-1252')

# Mudar SUM(IS_vIS) para excluir tipo=1 (IS_qUnid=0)
OLD_SQL = (
    '"ISNULL(SUM(IS_vBC),      0) AS vBCIS,         " & _\r\n'
    '       "ISNULL(SUM(IS_vIS),      0) AS vIS            " & _\r\n'
)
NEW_SQL = (
    '"ISNULL(SUM(IS_vBC),      0) AS vBCIS,         " & _\r\n'
    '       "ISNULL(SUM(CASE WHEN ISNULL(IS_qUnid,0) > 0 THEN ISNULL(IS_vIS,0) ELSE 0 END), 0) AS vIS " & _\r\n'
)
c3 = text2.count(OLD_SQL)
print(f'\nNFe_Completa patch 2 (AtualizarTotaisNota vIS): {c3} ocorrencias (esperado: 1)')
if c3 == 1:
    text2 = text2.replace(OLD_SQL, NEW_SQL)
    out2 = text2.encode('windows-1252')
    out2 = out2.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH2, 'wb').write(out2)
    print('NFe_Completa.frm: OK')
else:
    print('NFe_Completa.frm: ABORTADO')
    sys.exit(1)

print('\n=== RESUMO ===')
print('IS tipo=1 (ad valorem puro, IS_qUnid=0) excluido de:')
print('  - GerarItensImpostoIS (evita <qTrib>0.0000</qTrib> no schema)')
print('  - TotvIS (ISTot.vIS so inclui tipo=2/3)')
print('  - AtualizarTotaisNota (ValorNota = vProd - desc + IBS + CBS + IS_adrem)')
print('Resultado: schema valido + ISTot.vIS = per-item IS sum')
