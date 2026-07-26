path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

errors = []
R = '\r\n'

def sub(label, old, new, c):
    n = c.count(old)
    if n != 1:
        errors.append(f'{label}: count={n}')
        return c
    print(f'{label} OK')
    return c.replace(old, new, 1)

# ── 1: Remover CBSpAliq, IBSUFpAliq, IBSMunpAliq do SELECT de produtos ────────
content = sub('produtos SELECT remove CBS/IBS',
    'IBSCBSCST, CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISCST, ISpIS,',
    'IBSCBSCST, ISCST, ISpIS,',
    content)

# ── 2: Substituir as 3 linhas r("CBSpAliq/IBSUFpAliq/IBSMunpAliq") ─────────────
# por: CBS da empresa + IBS da cidade do cliente
old_block = (
    '     vCBSpAliq = Format(ValidateNull(r("CBSpAliq")), "##,##0.00")' + R +
    '     vIBSUFpAliq = Format(ValidateNull(r("IBSUFpAliq")), "##,##0.00")' + R +
    '     vIBSMunpAliq = Format(ValidateNull(r("IBSMunpAliq")), "##,##0.00")'
)

new_block = (
    '     ' + "' CBS: aliquota da tabela empresa (registro unico)" + R +
    '     Dim rEmp As ADODB.Recordset' + R +
    '     RsOpen rEmp, "SELECT CBSpAliq FROM empresa"' + R +
    '     If Not rEmp.BOF Then' + R +
    '        vCBSpAliq = Format(ValidateNull(rEmp("CBSpAliq")), "##,##0.00")' + R +
    '     Else' + R +
    '        vCBSpAliq = "0,00"' + R +
    '     End If' + R +
    '     If rEmp.State <> 0 Then rEmp.Close' + R +
    '     Set rEmp = Nothing' + R +
    R +
    '     ' + "' IBS: aliquotas da cidade do destinatario" + R +
    '     vIBSUFpAliq = "0,00"' + R +
    '     vIBSMunpAliq = "0,00"' + R +
    '     If Val(TxtCodCliente.Text) > 0 Then' + R +
    '        Dim rCli As ADODB.Recordset' + R +
    '        Dim lCodIBGE As Long' + R +
    '        lCodIBGE = 0' + R +
    '        RsOpen rCli, "SELECT CodigoIBGE FROM cliente WHERE CODIGO = " & Val(TxtCodCliente.Text)' + R +
    '        If Not rCli.BOF Then lCodIBGE = CLng(ValidateNull(rCli("CodigoIBGE")))' + R +
    '        If rCli.State <> 0 Then rCli.Close' + R +
    '        Set rCli = Nothing' + R +
    '        If lCodIBGE > 0 Then' + R +
    '           Dim rCid As ADODB.Recordset' + R +
    '           RsOpen rCid, "SELECT IBSUFpAliq, IBSMunpAliq FROM Cidade WHERE CodigoMunicipio = " & lCodIBGE' + R +
    '           If Not rCid.BOF Then' + R +
    '              vIBSUFpAliq = Format(ValidateNull(rCid("IBSUFpAliq")), "##,##0.00")' + R +
    '              vIBSMunpAliq = Format(ValidateNull(rCid("IBSMunpAliq")), "##,##0.00")' + R +
    '           End If' + R +
    '           If rCid.State <> 0 Then rCid.Close' + R +
    '           Set rCid = Nothing' + R +
    '        End If' + R +
    '     End If'
)

content = sub('CBS/IBS lookup empresa+cidade', old_block, new_block, content)

if errors:
    print('ERRORS:', errors)
else:
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
