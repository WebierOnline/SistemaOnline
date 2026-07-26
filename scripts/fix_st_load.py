data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Declarar variaveis ST junto com as demais variaveis de produto
old_dims = (
    b'Dim vICMSCST As String\r\n'
    b'Dim vICMSAliq As String\r\n'
    b'Dim vpRedBC As String\r\n'
)
new_dims = (
    b'Dim vICMSCST As String\r\n'
    b'Dim vICMSAliq As String\r\n'
    b'Dim vpRedBC As String\r\n'
    b'Dim vPMVAST As String\r\n'
    b'Dim vPICMSST As String\r\n'
    b'Dim vPRedBCST As String\r\n'
)
c = data.count(old_dims)
print(f'1. Dims ST: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# 2. Adicionar campos ST no SELECT de Aliquotas_Produto
old_select = (
    b'sSQL = "SELECT codigo, descricao, INF_ADICIONA, EAN, COD_BARRA, unid_medida, ncm, tamanho, REF, fabricante, CFOP, ICMSCST, ICMSAliq, pRedBC, piscst, pisAliq, cofinscst, cofinsAliq, ipicst, ipiAliq, cest, CASE WHEN abs(combustivel) = 1 THEN \'Combust\xedvel\' ELSE \'\' END as vTProduto FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"\r\n'
)
new_select = (
    b'sSQL = "SELECT codigo, descricao, INF_ADICIONA, EAN, COD_BARRA, unid_medida, ncm, tamanho, REF, fabricante, CFOP, ICMSCST, ICMSAliq, pRedBC, piscst, pisAliq, cofinscst, cofinsAliq, ipicst, ipiAliq, cest, pMVAST, pICMSST, pRedBCST, CASE WHEN abs(combustivel) = 1 THEN \'Combust\xedvel\' ELSE \'\' END as vTProduto FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"\r\n'
)
c = data.count(old_select)
print(f'2. SELECT produto: {c}')
if c == 1: data = data.replace(old_select, new_select)

# 3. Carregar variaveis ST apos vCEST
old_load = (
    b'     vIPICST = ValidateNull(r("ipicst"))\r\n'
    b'     vIPIALIQ = Format(ValidateNull(r("ipiAliq")), "##,##0.00")\r\n'
    b'     vCEST = ValidateNull(r("cest"))\r\n'
)
new_load = (
    b'     vIPICST = ValidateNull(r("ipicst"))\r\n'
    b'     vIPIALIQ = Format(ValidateNull(r("ipiAliq")), "##,##0.00")\r\n'
    b'     vCEST = ValidateNull(r("cest"))\r\n'
    b'     vPMVAST = Format(ValidateNull(r("pMVAST")), "##,##0.00")\r\n'
    b'     vPICMSST = Format(ValidateNull(r("pICMSST")), "##,##0.00")\r\n'
    b'     vPRedBCST = Format(ValidateNull(r("pRedBCST")), "##,##0.00")\r\n'
)
c = data.count(old_load)
print(f'3. Load variaveis ST: {c}')
if c == 1: data = data.replace(old_load, new_load)

# 4. Zerar variaveis ST no LimparObjetosProduto / reset
old_reset = (
    b'     vICMSCST = ""\r\n'
    b'     vICMSAliq = ""\r\n'
)
new_reset = (
    b'     vICMSCST = ""\r\n'
    b'     vICMSAliq = ""\r\n'
    b'     vPMVAST = ""\r\n'
    b'     vPICMSST = ""\r\n'
    b'     vPRedBCST = ""\r\n'
)
c = data.count(old_reset)
print(f'4. Reset variaveis ST: {c}')
if c >= 1: data = data.replace(old_reset, new_reset)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
