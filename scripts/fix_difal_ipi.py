data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Declarar vIPICompoeDIFAL junto com vRegimeTributario
old_dim = b'Dim vRegimeTributario As Integer\r\n'
new_dim = (
    b'Dim vRegimeTributario As Integer\r\n'
    b'Dim vIPICompoeDIFAL As Integer\r\n'
)
c = data.count(old_dim)
print(f'1. Dim: {c}')
if c == 1: data = data.replace(old_dim, new_dim)

# 2. Adicionar IPICompoeDIFAL no SELECT e no carregamento (ocorre 2 vezes identicas)
old_select = (
    b'sSQL = "SELECT CRT, ESTADO, RegimeTributario FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
    b'    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))\r\n'
)
new_select = (
    b'sSQL = "SELECT CRT, ESTADO, RegimeTributario, IPICompoeDIFAL FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
    b'    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))\r\n'
    b'    vIPICompoeDIFAL = IIf(IsNull(r("IPICompoeDIFAL")), 0, r("IPICompoeDIFAL"))\r\n'
)
c = data.count(old_select)
print(f'2. SELECT Empresa (esperado 2): {c}')
if c >= 1: data = data.replace(old_select, new_select)

# 3. Ajustar base do DIFAL para incluir IPI quando parametro ativo
old_difal_base = (
    b"            ' 3. Base de calculo (base dupla ou simples)\r\n"
    b'            vBaseItem = CDbl(txtSubTotal.Text)\r\n'
)
new_difal_base = (
    b"            ' 3. Base de calculo (base dupla ou simples)\r\n"
    b'            vBaseItem = CDbl(txtSubTotal.Text)\r\n'
    b"            ' Inclui IPI na base do DIFAL se parametro ativo (Art. 13 Lei Kandir - consumidor final)\r\n"
    b'            If vIPICompoeDIFAL = 1 Then vBaseItem = vBaseItem + CDbl(vValorIPI)\r\n'
)
c = data.count(old_difal_base)
print(f'3. Base DIFAL + IPI: {c}')
if c == 1: data = data.replace(old_difal_base, new_difal_base)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
