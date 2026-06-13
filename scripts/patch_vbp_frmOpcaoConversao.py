data = open(r'C:\Projeto\OnlineCommerce\OnlineCommerce.vbp', 'rb').read()
text = data.decode('windows-1252')

old = 'Form=Forms\\NFe_Completa.frm\r\n'
new = 'Form=Forms\\NFe_Completa.frm\r\nForm=Forms\\frmOpcaoConversao.frm\r\n'

c = text.count(old)
print('count:', c)
if c == 1:
    text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(r'C:\Projeto\OnlineCommerce\OnlineCommerce.vbp', 'wb').write(out)
    print('OK')
else:
    print('ERRO')
