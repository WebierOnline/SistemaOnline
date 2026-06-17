# -*- coding: utf-8 -*-
# Restaura as chamadas GerarItensImpostoIS em modNFe.bas com parametros originais.
# DLL foi atualizada: agora gera <oISEspec> (valido) em vez de <eTrib> (invalido).
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = (
    "           'GerarItensImpostoIS desabilitado: DLL gera <eTrib> (invalido no schema NT 2025.003)\r\n"
    "           'If vIS > 0 Then\r\n"
    "           '   iRetorno = sistNFe.GerarItensImpostoIS(...)\r\n"
    "           'End If\r\n"
)
NEW = (
    '           If vIS > 0 Then\r\n'
    '              iRetorno = sistNFe.GerarItensImpostoIS('
        'IIf(IsNull(NFeItens!IS_CST), "99", NFeItens!IS_CST), '
        'IIf(IsNull(NFeItens!cClassTrib_IS), "", NFeItens!cClassTrib_IS), '
        'vBCIS, pIS, dISvUnid, '
        'IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), '
        'dISqUnid, vIS, mensagemAlerta, mensagemErro)\r\n'
    '           End If\r\n'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — GerarItensImpostoIS restaurado com parametros originais')
else:
    print('ABORTADO')
