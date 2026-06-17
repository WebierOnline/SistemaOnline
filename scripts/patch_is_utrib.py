# -*- coding: utf-8 -*-
# Corrige qTrib=0: quando dISqUnid=0, passa uTrib="" para a DLL nao gerar
# estrutura por unidade (<uTrib>/<qTrib>) com quantidade zero.
PATH = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

OLD = (
    'vBCIS, pIS, dISvUnid, '
    'IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), '
    'dISqUnid, vIS, mensagemAlerta, mensagemErro)'
)
NEW = (
    'vBCIS, pIS, dISvUnid, '
    'IIf(dISqUnid > 0, IIf(IsNull(NFeItens!UnidadeTributavel), "", NFeItens!UnidadeTributavel), ""), '
    'dISqUnid, vIS, mensagemAlerta, mensagemErro)'
)

count = text.count(OLD)
print(f'Ocorrencias encontradas: {count} (esperado: 2)')

if count == 2:
    text = text.replace(OLD, NEW)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('OK — uTrib zerado quando dISqUnid = 0')
else:
    print('ABORTADO')
