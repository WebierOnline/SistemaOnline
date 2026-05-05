# -*- coding: utf-8 -*-
FRM = 'C:/Projeto/Compartilhado/Forms/Produtos_Cadastro.frm'
data = open(FRM, 'rb').read()
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')

# Move TABACARIA do case padrao para o case 000+IS (junto com bebidas alcoolicas)
old = (b'        Case \"BEBIDAS (ALCO\xd3LICAS)\", \"BEBIDAS (A\xc7UCARADAS)\"\r\n'
       b'            sIBSCST = \"000\": sCClassTrib = \"000001\"')
new  = (b'        Case \"BEBIDAS (ALCO\xd3LICAS)\", \"BEBIDAS (A\xc7UCARADAS)\", \"TABACARIA\"\r\n'
        b'            sIBSCST = \"000\": sCClassTrib = \"000001\"')
subs = [('case 000+IS adiciona TABACARIA', old, new)]

old2 = (b'             \"LATIC\xcdNIOS\", \"CONGELADOS\", \"TABACARIA\"')
new2 = (b'             \"LATIC\xcdNIOS\", \"CONGELADOS\"')
subs.append(('case padrao remove TABACARIA', old2, new2))

ok = 0
for name, o, n in subs:
    o_n = o.replace(b'\r\n', b'\n')
    d_n = data.replace(b'\r\n', b'\n')
    count = d_n.count(o_n)
    if count == 0: print(f'NENHUMA: {name}'); continue
    if count > 1: print(f'AMBIGUO: {name}'); continue
    data = d_n.replace(o_n, n.replace(b'\r\n', b'\n'))
    print(f'OK: {name}'); ok += 1

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(FRM, 'wb').write(data)
print(f'\n{ok}/{len(subs)} substituicoes aplicadas.')
