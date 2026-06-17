# -*- coding: utf-8 -*-
# Corrige qTrib=0 quando IS e calculado por unidade (dISvUnid > 0).
# Se IS_qUnid nao foi gravado no banco, recalcula: qTrib = vIS / dISvUnid.
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

# Insere o calculo de dISqUnid derivado logo apos a leitura das variaveis IS,
# antes do If vIS > 0. Aplica nas duas ocorrencias (mesmo bloco OLD).
OLD = (
    '           dISqUnid = IIf(IsNull(NFeItens!IS_qUnid), 0, CDbl(NFeItens!IS_qUnid))\r\n'
    '           dISvUnid = IIf(IsNull(NFeItens!IS_vUnid), 0, CDbl(NFeItens!IS_vUnid))\r\n'
    '           If vIS > 0 Then\r\n'
)
NEW = (
    '           dISqUnid = IIf(IsNull(NFeItens!IS_qUnid), 0, CDbl(NFeItens!IS_qUnid))\r\n'
    '           dISvUnid = IIf(IsNull(NFeItens!IS_vUnid), 0, CDbl(NFeItens!IS_vUnid))\r\n'
    '           If dISvUnid > 0 And dISqUnid = 0 And vIS > 0 Then dISqUnid = vIS / dISvUnid\r\n'
    '           If vIS > 0 Then\r\n'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — dISqUnid calculado automaticamente quando dISvUnid > 0 e qUnid = 0')
else:
    print('ABORTADO')
