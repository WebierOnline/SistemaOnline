data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# 1. Adicionar Dims para PIS, COFINS no AtualizarTotaisNota
old_dims = (
    b'Dim varICMSUFDest As Double\r\n'
    b'Dim varFCPUFDest  As Double\r\n'
    b'Dim varNota       As Double\r\n'
)
new_dims = (
    b'Dim varICMSUFDest As Double\r\n'
    b'Dim varFCPUFDest  As Double\r\n'
    b'Dim varPIS        As Double\r\n'
    b'Dim varCOFINS     As Double\r\n'
    b'Dim varNota       As Double\r\n'
)
c = data.count(old_dims)
print(f'1. Dims PIS/COFINS/ST: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# 2. Remover leitura manual de ST e substituir SELECT para incluir ST e PIS/COFINS dos itens
old_select = (
    b"' ICMS-ST e BaseICMSST s\xe3o informados manualmente no cabe\xe7alho\r\n"
    b'varICMSST = IIf(Vazio(txtValorICMSST), 0, CDbl(Format(txtValorICMSST, "##0.00")))\r\n'
    b'varBaseICMSST = IIf(Vazio(txtBaseICMSST), 0, CDbl(Format(txtBaseICMSST, "##0.00")))\r\n'
    b'\r\n'
    b"' Todos os demais totais v\xeam dos itens\r\n"
    b'sSQL = "SELECT " & _\r\n'
    b'       "ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS ValorProdutos, " & _\r\n'
    b'       "ISNULL(SUM(ValorFrete),   0) AS ValorFrete,   " & _\r\n'
    b'       "ISNULL(SUM(ValorSeguro),  0) AS ValorSeguro,  " & _\r\n'
    b'       "ISNULL(SUM(ValorOutros),  0) AS ValorOutros,  " & _\r\n'
    b'       "ISNULL(SUM(ValorDesconto),0) AS ValorDesconto," & _\r\n'
    b'       "ISNULL(SUM(IPIvIPI),      0) AS ValorIPI,     " & _\r\n'
    b'       "ISNULL(SUM(vICMS),        0) AS ValorICMS,    " & _\r\n'
    b'       "ISNULL(SUM(vBC),          0) AS BaseICMS,     " & _\r\n'
    b'       "ISNULL(SUM(vICMSUFDest),  0) AS vICMSUFDest, " & _\r\n'
    b'       "ISNULL(SUM(vFCPUFDest),   0) AS vFCPUFDest   " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
)
new_select = (
    b"' Todos os totais vem dos itens\r\n"
    b'sSQL = "SELECT " & _\r\n'
    b'       "ISNULL(SUM(ValorUnitarioComercializacao * QuantidadeComercial), 0) AS ValorProdutos, " & _\r\n'
    b'       "ISNULL(SUM(ValorFrete),   0) AS ValorFrete,   " & _\r\n'
    b'       "ISNULL(SUM(ValorSeguro),  0) AS ValorSeguro,  " & _\r\n'
    b'       "ISNULL(SUM(ValorOutros),  0) AS ValorOutros,  " & _\r\n'
    b'       "ISNULL(SUM(ValorDesconto),0) AS ValorDesconto," & _\r\n'
    b'       "ISNULL(SUM(IPIvIPI),      0) AS ValorIPI,     " & _\r\n'
    b'       "ISNULL(SUM(vICMS),        0) AS ValorICMS,    " & _\r\n'
    b'       "ISNULL(SUM(vBC),          0) AS BaseICMS,     " & _\r\n'
    b'       "ISNULL(SUM(vBCST),        0) AS BaseICMSST,   " & _\r\n'
    b'       "ISNULL(SUM(vICMSST),      0) AS ValorICMSST,  " & _\r\n'
    b'       "ISNULL(SUM(PISvPIS),      0) AS ValorPIS,     " & _\r\n'
    b'       "ISNULL(SUM(cofinsvcofins),0) AS ValorCOFINS,  " & _\r\n'
    b'       "ISNULL(SUM(vICMSUFDest),  0) AS vICMSUFDest,  " & _\r\n'
    b'       "ISNULL(SUM(vFCPUFDest),   0) AS vFCPUFDest    " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
)
c = data.count(old_select)
print(f'2. SELECT totais: {c}')
if c == 1: data = data.replace(old_select, new_select)

# 3. Adicionar leitura de ST, PIS, COFINS no bloco If Not rTotais.EOF
old_read = (
    b'    varICMSUFDest = ValidateNull(rTotais("vICMSUFDest"))\r\n'
    b'    varFCPUFDest = ValidateNull(rTotais("vFCPUFDest"))\r\n'
    b'End If\r\n'
)
new_read = (
    b'    varICMSUFDest = ValidateNull(rTotais("vICMSUFDest"))\r\n'
    b'    varFCPUFDest = ValidateNull(rTotais("vFCPUFDest"))\r\n'
    b'    varICMSST = ValidateNull(rTotais("ValorICMSST"))\r\n'
    b'    varBaseICMSST = ValidateNull(rTotais("BaseICMSST"))\r\n'
    b'    varPIS = ValidateNull(rTotais("ValorPIS"))\r\n'
    b'    varCOFINS = ValidateNull(rTotais("ValorCOFINS"))\r\n'
    b'End If\r\n'
)
c = data.count(old_read)
print(f'3. Leitura ST/PIS/COFINS: {c}')
if c == 1: data = data.replace(old_read, new_read)

# 4. Atualizar textboxes: adicionar txtValorICMSST e txtBaseICMSST
old_txt = (
    b'txtValorICMS.Text = FormatNumber(varICMS, 2)\r\n'
    b'txtBaseICMS.Text = FormatNumber(varBaseICMS, 2)\r\n'
    b'txtTotaldaNota.Text = FormatNumber(varNota, 2)\r\n'
)
new_txt = (
    b'txtValorICMS.Text = FormatNumber(varICMS, 2)\r\n'
    b'txtBaseICMS.Text = FormatNumber(varBaseICMS, 2)\r\n'
    b'txtValorICMSST.Text = FormatNumber(varICMSST, 2)\r\n'
    b'txtBaseICMSST.Text = FormatNumber(varBaseICMSST, 2)\r\n'
    b'txtTotaldaNota.Text = FormatNumber(varNota, 2)\r\n'
)
c = data.count(old_txt)
print(f'4. Textboxes ST: {c}')
if c == 1: data = data.replace(old_txt, new_txt)

# 5. Adicionar ValorPIS e ValorCOFINS no UPDATE NotaFiscal
old_update = (
    b'       "ValorICMSST         = " & FSQL(varICMSST, 2) & ", " & _\r\n'
    b'       "BaseICMSST          = " & FSQL(varBaseICMSST, 2) & ", " & _\r\n'
    b'       "vICMSUFDest         = " & FSQL(varICMSUFDest, 2) & ", " & _\r\n'
)
new_update = (
    b'       "ValorICMSST         = " & FSQL(varICMSST, 2) & ", " & _\r\n'
    b'       "BaseICMSST          = " & FSQL(varBaseICMSST, 2) & ", " & _\r\n'
    b'       "ValorPIS            = " & FSQL(varPIS, 2) & ", " & _\r\n'
    b'       "ValorCOFINS         = " & FSQL(varCOFINS, 2) & ", " & _\r\n'
    b'       "vICMSUFDest         = " & FSQL(varICMSUFDest, 2) & ", " & _\r\n'
)
c = data.count(old_update)
print(f'5. UPDATE PIS/COFINS: {c}')
if c == 1: data = data.replace(old_update, new_update)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
