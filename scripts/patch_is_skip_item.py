# -*- coding: utf-8 -*-
# Desabilita chamadas a GerarItensImpostoIS em modNFe.bas.
# Motivo: DLL antiga gera <eTrib> como container, elemento rejeitado pelo schema
# atual da NT 2025.003 (espera oISEspec ou uTrib/qTrib diretamente).
# TotvIS continua acumulado normalmente e vai para ISTot.vIS via GerarTotalIBSCBS.
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

# As duas ocorrencias sao identicas apos o patch anterior (dISvUnid e dISqUnid ja zerados)
OLD = (
    '           If vIS > 0 Then\r\n'
    '              iRetorno = sistNFe.GerarItensImpostoIS(IIf(IsNull(NFeItens!IS_CST), "99", NFeItens!IS_CST), IIf(IsNull(NFeItens!cClassTrib_IS), "", NFeItens!cClassTrib_IS), vBCIS, pIS, 0, "", 0, vIS, mensagemAlerta, mensagemErro)\r\n'
    '           End If\r\n'
)
NEW = (
    '           ' + "'" + 'GerarItensImpostoIS desabilitado: DLL gera <eTrib> (invalido no schema NT 2025.003)\r\n'
    "           'If vIS > 0 Then\r\n"
    "           '   iRetorno = sistNFe.GerarItensImpostoIS(...)\r\n"
    "           'End If\r\n"
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — GerarItensImpostoIS desabilitado nas duas ocorrencias')
else:
    print('ABORTADO')
