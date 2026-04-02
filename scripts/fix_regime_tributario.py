data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Declarar vRegimeTributario junto com vTipoCRT
old1 = b'Dim vTipoCRT As Integer\r\n'
new1 = (
    b'Dim vTipoCRT As Integer\r\n'
    b'Dim vRegimeTributario As Integer\r\n'
)
c = data.count(old1)
print(f'1. Dim: {c}')
if c == 1: data = data.replace(old1, new1)

# 2. Ambos os SELECTs + carregamentos (sao identicos, substituir os dois)
old2 = (
    b'sSQL = "SELECT CRT, ESTADO FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
)
new2 = (
    b'sSQL = "SELECT CRT, ESTADO, RegimeTributario FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
    b'    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))\r\n'
)
c = data.count(old2)
print(f'2. SELECT + load (esperado 2): {c}')
if c >= 1: data = data.replace(old2, new2)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
