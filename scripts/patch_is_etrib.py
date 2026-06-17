# -*- coding: utf-8 -*-
# Corrige GerarItensImpostoIS em modNFe.bas:
# Passa dISvUnid=0 e dISqUnid=0 em vez dos valores do DB.
# Motivo: DLL gera <eTrib> quando dISvUnid>0, elemento rejeitado pelo schema atual
# (NT 2025.003 espera oISEspec, uTrib ou vIS direto — nunca eTrib).
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = (
    'vBCIS, pIS, dISvUnid, '
    'IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), '
    'dISqUnid, vIS, mensagemAlerta, mensagemErro)'
)
NEW = (
    'vBCIS, pIS, 0, '
    '"", '
    '0, vIS, mensagemAlerta, mensagemErro)'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — dISvUnid e dISqUnid zerados nas duas chamadas GerarItensImpostoIS')
else:
    print('ABORTADO — contagem inesperada, arquivo nao modificado')
