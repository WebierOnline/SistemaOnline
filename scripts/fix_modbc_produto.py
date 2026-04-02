data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Declarar vModBC junto com as demais variaveis de produto
old_dim = b'Dim vPMVAST As String\r\n'
new_dim = (
    b'Dim vModBC As String\r\n'
    b'Dim vPMVAST As String\r\n'
)
c = data.count(old_dim)
print(f'1. Dim vModBC: {c}')
if c == 1: data = data.replace(old_dim, new_dim)

# 2. Adicionar modBC no SELECT de Aliquotas_Produto
old_select = (
    b'sSQL = "SELECT codigo, descricao, INF_ADICIONA, EAN, COD_BARRA, unid_medida, ncm, tamanho, REF, fabricante, CFOP, ICMSCST, ICMSAliq, pRedBC, piscst, pisAliq, cofinscst, cofinsAliq, ipicst, ipiAliq, cest, pMVAST, pICMSST, pRedBCST, CASE WHEN abs(combustivel) = 1 THEN \'Combust\xedvel\' ELSE \'\' END as vTProduto FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"\r\n'
)
new_select = (
    b'sSQL = "SELECT codigo, descricao, INF_ADICIONA, EAN, COD_BARRA, unid_medida, ncm, tamanho, REF, fabricante, CFOP, ICMSCST, ICMSAliq, pRedBC, modBC, piscst, pisAliq, cofinscst, cofinsAliq, ipicst, ipiAliq, cest, pMVAST, pICMSST, pRedBCST, CASE WHEN abs(combustivel) = 1 THEN \'Combust\xedvel\' ELSE \'\' END as vTProduto FROM produtos WHERE (codigo = " & txtCodProduto.Text & ");"\r\n'
)
c = data.count(old_select)
print(f'2. SELECT produto: {c}')
if c == 1: data = data.replace(old_select, new_select)

# 3. Carregar vModBC apos vCEST
old_load = (
    b'     vCEST = ValidateNull(r("cest"))\r\n'
    b'     vPMVAST = Format(ValidateNull(r("pMVAST")), "##,##0.00")\r\n'
)
new_load = (
    b'     vCEST = ValidateNull(r("cest"))\r\n'
    b'     vModBC = ValidateNull(r("modBC"))\r\n'
    b'     vPMVAST = Format(ValidateNull(r("pMVAST")), "##,##0.00")\r\n'
)
c = data.count(old_load)
print(f'3. Load vModBC: {c}')
if c == 1: data = data.replace(old_load, new_load)

# 4. Zerar vModBC no reset
old_reset = (
    b'     vPMVAST = ""\r\n'
    b'     vPICMSST = ""\r\n'
    b'     vPRedBCST = ""\r\n'
)
new_reset = (
    b'     vModBC = ""\r\n'
    b'     vPMVAST = ""\r\n'
    b'     vPICMSST = ""\r\n'
    b'     vPRedBCST = ""\r\n'
)
c = data.count(old_reset)
print(f'4. Reset vModBC: {c}')
if c >= 1: data = data.replace(old_reset, new_reset)

# 5. Usar vModBC em Load_Data_Itens (substituir o valor fixo 3)
old_modbc = b'    Tb("modBC") = Format(3, "@")\r\n'
new_modbc = b'    Tb("modBC") = Format(IIf(vModBC = "", 3, vModBC), "@")\r\n'
c = data.count(old_modbc)
print(f'5. Tb modBC dinamico: {c}')
if c == 1: data = data.replace(old_modbc, new_modbc)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
