# -*- coding: utf-8 -*-
# Corrige geracao do bloco IS por item: chama GerarItensImpostoIS somente quando
# dISqUnid > 0 (IS ad rem, tipo 2). Para IS ad valorem (tipo 1, dISqUnid=0), a DLL
# gera <qTrib>0.0000</qTrib> que o schema rejeita. O TotvIS acumula normalmente
# para ISTot em ambos os casos.
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = '           If vIS > 0 Then\r\n'
NEW = '           If vIS > 0 And dISqUnid > 0 Then\r\n'

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — GerarItensImpostoIS chamado apenas quando dISqUnid > 0')
else:
    print('ABORTADO')
