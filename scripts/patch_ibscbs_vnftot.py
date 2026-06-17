# -*- coding: utf-8 -*-
# Corrige vNFTot em GerarTotalIBSCBS em modNFe.bas:
# A formula anterior (ValorProdutos - ValorDesconto + IBS + CBS + IS) nao incluia
# frete/seguro/outros/IPI, causando IBSCBSTot.vNFTot != ICMSTot.vNF quando
# a nota possui algum desses valores. Substitui pela NFe!ValorNota, que ja
# contem todos os componentes (calculado por AtualizarTotaisNota antes da transmissao).
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = (
    'GerarTotalIBSCBS(TotvIS, TotvBCCBSIBS, 0, 0, TotvIBSUF, 0, 0, TotvIBSMun, TotvIBS, '
    '0, 0, 0, 0, TotvCBS, 0, 0, 0, 0, 0, 0, 0, 0, '
    '(NFe!ValorProdutos - NFe!ValorDesconto + TotvIBS + TotvCBS + TotvIS), '
    'mensagemAlerta, mensagemErro)'
)

NEW = (
    'GerarTotalIBSCBS(TotvIS, TotvBCCBSIBS, 0, 0, TotvIBSUF, 0, 0, TotvIBSMun, TotvIBS, '
    '0, 0, 0, 0, TotvCBS, 0, 0, 0, 0, 0, 0, 0, 0, '
    'NFe!ValorNota, '
    'mensagemAlerta, mensagemErro)'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 1)')

if count == 1:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — IBSCBSTot.vNFTot agora usa NFe!ValorNota (igual a ICMSTot.vNF)')
else:
    print('ABORTADO')
