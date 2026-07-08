"""
Patch: Vendas_Consulta_PorProdutos.frm
Adiciona filtro AND OS.STATUS_OS = 1 no bloco POR SERVICOS,
excluindo OS abertas (EM EXECUCAO / A COMECAR) das consultas de servicos.

Por que: OS com DATA_TERMINO no mes mas STATUS_OS=0 inflam o total
de servicos (OS 438 = R$750 extra em Abril/2026).
"""

FRM = r'C:\Projeto\OnlineCommerce\Forms\Vendas_Consulta_PorProdutos.frm'

with open(FRM, 'rb') as f:
    raw = f.read()

# Normaliza para LF para trabalhar como texto
data = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('windows-1252')

patches = []

# 1. Embed WHERE OS.STATUS_OS = 1 no sBase do POR SERVICOS
patches.append((
    '"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS "',
    '"FROM OS_Servicos_Auto s INNER JOIN OS ON s.cod_os = OS.COD_OS WHERE OS.STATUS_OS = 1 "',
    1
))

# 2. Todas as 13 ocorrencias de sBase & "WHERE no bloco POR SERVICOS
#    passam a ser AND (sBase ja tem o WHERE embutido)
patches.append((
    'sBase & "WHERE ',
    'sBase & "AND ',
    13
))

ok = True
for old, new, expected in patches:
    count = text.count(old)
    if count != expected:
        print(f'ERRO [{repr(old[:40])}]: {count} ocorrencias (esperado {expected})')
        ok = False
    else:
        print(f'OK [{repr(old[:40])}]: {count} ocorrencia(s)')

if not ok:
    print('Nenhuma alteracao feita.')
else:
    for old, new, _ in patches:
        text = text.replace(old, new)

    result = text.encode('windows-1252').replace(b'\n', b'\r\n')
    with open(FRM, 'wb') as f:
        f.write(result)
    print('Arquivo salvo com sucesso.')
