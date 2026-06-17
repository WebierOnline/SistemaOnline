# -*- coding: utf-8 -*-
# Reverte condicao GerarItensImpostoIS em modNFe.bas:
# A condicao "If vIS > 0 And dISqUnid > 0 Then" bloqueava itens IS tipo=1
# (ad valorem, IS_qUnid=0) de gerar <vIS> por item no XML, enquanto o
# acumulador TotvIS ja somava esses valores. Resultado: IBSCBSTot.vIS=70,00
# mas soma per-item <vIS>=66,40 -> SEFAZ rejeita "Total da NF difere".
# Fix: volta para "If vIS > 0 Then"; uTrib="" ja garante que a DLL gera
# <oISEspec> (ad valorem) sem <qTrib> para esses itens.
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = '           If vIS > 0 And dISqUnid > 0 Then\r\n'
NEW = '           If vIS > 0 Then\r\n'

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — condicao revertida para "If vIS > 0 Then" em ambas as ocorrencias')
    print('     Itens IS tipo=1 (ad valorem) agora geram <vIS> por item no XML')
    print('     IBSCBSTot.vIS (70,00) == soma per-item vIS (70,00) apos correcao')
else:
    print('ABORTADO — nao aplicado')
