data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old = (
    b'    End If\r\n'
    b'    \r\n'
    b"    'PIS e COFINS\r\n"
)
new = (
    b'    End If\r\n'
    b'    If CDbl(Tb("vBC")) = 0 Then Tb("modBC") = ""\r\n'
    b'    \r\n'
    b"    'PIS e COFINS\r\n"
)

c = data.count(old)
print(f'count: {c}')
if c == 1:
    data = data.replace(old, new)
    data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
    print('Salvo.')
