path = r'C:\Projeto\Compartilhado\Modulos\modNFe.bas'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c, expected=1):
    n = c.count(old)
    if n != expected:
        errors.append(f'{label}: count={n} (esperado {expected})')
        return c
    print(f'{label} OK (n={n})')
    return c.replace(old, new)

# 1 — Adicionar Dims
content = sub('Dim pAliqEfet*',
    'Dim pIBSpRed As Double, pCBSpRed As Double',
    'Dim pIBSpRed As Double, pCBSpRed As Double' + R +
    'Dim pAliqEfetUF As Double, pAliqEfetMun As Double, pAliqEfetCBS As Double',
    content)

# 2 — Calcular pAliqEfet* apos leitura de pCBSpRed (aparece 2x — ambas as branches)
content = sub('Calcular pAliqEfet* (2x)',
    '           pCBSpRed = IIf(IsNull(NFeItens!CBS_pRed), 0, CDbl(NFeItens!CBS_pRed))' + R +
    '           If vBCCBSIBS > 0 Then',

    '           pCBSpRed = IIf(IsNull(NFeItens!CBS_pRed), 0, CDbl(NFeItens!CBS_pRed))' + R +
    '           pAliqEfetUF  = pIBSUF  * (1 - pIBSpRed  / 100)' + R +
    '           pAliqEfetMun = pIBSMun * (1 - pIBSpRed  / 100)' + R +
    '           pAliqEfetCBS = pCBS    * (1 - pCBSpRed  / 100)' + R +
    '           If vBCCBSIBS > 0 Then',
    content, expected=2)

# 3 — Atualizar chamada: substituir pIBSUF/pIBSMun/pCBS nos slots pAliqEfet (2x)
content = sub('pAliqEfet na chamada (2x)',
    '                  pIBSUF, 0, 0, 0, pIBSpRed, pIBSUF, vIBSUF, _' + R +
    '                  pIBSMun, 0, 0, 0, pIBSpRed, pIBSMun, vIBSMun, vIBS, _' + R +
    '                  pCBS, 0, 0, 0, pCBSpRed, pCBS, vCBS, _',

    '                  pIBSUF, 0, 0, 0, pIBSpRed, pAliqEfetUF, vIBSUF, _' + R +
    '                  pIBSMun, 0, 0, 0, pIBSpRed, pAliqEfetMun, vIBSMun, vIBS, _' + R +
    '                  pCBS, 0, 0, 0, pCBSpRed, pAliqEfetCBS, vCBS, _',
    content, expected=2)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
