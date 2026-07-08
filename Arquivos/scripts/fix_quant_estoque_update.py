path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')

old = (
"dbData.Execute \"UPDATE produtos SET quant_estoque = quant_estoque + \" & Replace(txtQuantFinal.Text, \",\", \".\") & \" WHERE (codigo = \" & txtCodProdExist.Text & \");\"\r\n"
)

new = (
"dbData.Execute \"UPDATE produtos SET quant_estoque = quant_estoque + \" & Replace(CDbl(txtQuantFinal.Text), \",\", \".\") & \" WHERE (codigo = \" & txtCodProdExist.Text & \");\"\r\n"
)

print('found:', data.count(old))
data2 = data.replace(old, new, 1)
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
