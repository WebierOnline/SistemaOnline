path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

old = (
    "    ' Ad Valorem: qUnid e vUnid nao se aplicam (zerar para evitar rejeicao SEFAZ)" + R +
    "    If iTipoIS = 1 Then" + R +
    "        curISqUnid = 0" + R +
    "        curISvUnid = 0" + R +
    "    End If"
)

new = (
    "    ' Ad Valorem e sem IS: qUnid e vUnid nao se aplicam (zerar para evitar rejeicao SEFAZ)" + R +
    "    If iTipoIS = 0 Or iTipoIS = 1 Then" + R +
    "        curISqUnid = 0" + R +
    "        curISvUnid = 0" + R +
    "    End If"
)

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
