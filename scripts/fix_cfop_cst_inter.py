data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# =============================================================================
# ALTERACAO 1: Load_Data_Itens - ajustar CFOP (5->6 se inter) e CST (por regime+CFOP)
# =============================================================================
old_cfop_cst = (
    b'    Tb("CFOP") = Format(vCFOP, "@")\r\n'
    b'    Tb("NCM") = Format(vNCM, "@")\r\n'
    b'    Tb("UnidadeComercial") = UCase(Format(vUnid_medida, "@"))\r\n'
    b'    vValorProdutos = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'    \r\n'
    b'    \'ICMS\r\n'
    b'    Tb("CST") = Right(Format(vICMSCST, "@"), 3)\r\n'
)
new_cfop_cst = (
    b'    \' CFOP: converter 5xxx -> 6xxx se operacao interestadual\r\n'
    b'    Dim sCFOPFinal As String\r\n'
    b'    sCFOPFinal = Format(vCFOP, "@")\r\n'
    b'    If Left(cboDestOperacao.Text, 1) = "2" Then\r\n'
    b'        If Left(sCFOPFinal, 1) = "5" Then sCFOPFinal = "6" & Mid(sCFOPFinal, 2)\r\n'
    b'    End If\r\n'
    b'    Tb("CFOP") = sCFOPFinal\r\n'
    b'    Tb("NCM") = Format(vNCM, "@")\r\n'
    b'    Tb("UnidadeComercial") = UCase(Format(vUnid_medida, "@"))\r\n'
    b'    vValorProdutos = CDbl(Format(txtSubTotal, "@"))\r\n'
    b'    \r\n'
    b'    \'ICMS\r\n'
    b'    \' CST/CSOSN: para Simples (1,2), derivar do CFOP original do cadastro\r\n'
    b'    Dim sCSTFinal As String\r\n'
    b'    If vRegimeTributario = 1 Or vRegimeTributario = 2 Then\r\n'
    b'        If Right(Format(vCFOP, "@"), 3) = "102" Then\r\n'
    b'            sCSTFinal = "102"\r\n'
    b'        ElseIf Right(Format(vCFOP, "@"), 3) = "405" Then\r\n'
    b'            sCSTFinal = "500"\r\n'
    b'        Else\r\n'
    b'            sCSTFinal = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'        End If\r\n'
    b'    Else\r\n'
    b'        sCSTFinal = Right(Format(vICMSCST, "@"), 3)\r\n'
    b'    End If\r\n'
    b'    vICMSCST = sCSTFinal\r\n'
    b'    Tb("CST") = sCSTFinal\r\n'
)
c = data.count(old_cfop_cst)
print(f'1. CFOP/CST em Load_Data_Itens: {c}')
if c == 1: data = data.replace(old_cfop_cst, new_cfop_cst)

# =============================================================================
# ALTERACAO 2a: Inserir sub AtualizarCFOPCSTItens antes de cboDestOperacao_Change
# =============================================================================
old_change = b'Private Sub cboDestOperacao_Change()\r\n'
new_change = (
    b'Private Sub AtualizarCFOPCSTItens()\r\n'
    b'    If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'    Dim vCodNota As Long\r\n'
    b'    vCodNota = Val(txtCodNota.Text)\r\n'
    b'    Dim bInter As Boolean\r\n'
    b'    bInter = (Left(cboDestOperacao.Text, 1) = "2")\r\n'
    b'\r\n'
    b'    \' 1. Converter CFOP\r\n'
    b'    If bInter Then\r\n'
    b'        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = \'6\' + SUBSTRING(CFOP, 2, 3) WHERE CodigoNota = " & vCodNota & " AND LEFT(CFOP, 1) = \'5\'"\r\n'
    b'    Else\r\n'
    b'        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = \'5\' + SUBSTRING(CFOP, 2, 3) WHERE CodigoNota = " & vCodNota & " AND LEFT(CFOP, 1) = \'6\'"\r\n'
    b'    End If\r\n'
    b'\r\n'
    b'    \' 2. Atualizar CST/CSOSN para Simples (regime 1 ou 2)\r\n'
    b'    If vRegimeTributario = 1 Or vRegimeTributario = 2 Then\r\n'
    b'        dbData.Execute "UPDATE NotaFiscalItens SET CST = CASE WHEN RIGHT(CFOP, 3) = \'102\' THEN \'102\' WHEN RIGHT(CFOP, 3) = \'405\' THEN \'500\' ELSE CST END WHERE CodigoNota = " & vCodNota\r\n'
    b'    End If\r\n'
    b'\r\n'
    b'    \' 3. Recalcular impostos e exibir grid\r\n'
    b'    RecalcularItensNota\r\n'
    b'    Exibir_Itens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboDestOperacao_Change()\r\n'
)
c = data.count(old_change)
print(f'2a. Inserir AtualizarCFOPCSTItens: {c}')
if c == 1: data = data.replace(old_change, new_change)

# =============================================================================
# ALTERACAO 2b: chamar AtualizarCFOPCSTItens ao fim de cboDestOperacao_LostFocus
# =============================================================================
old_lostfocus_end = (
    b'If cboDestOperacao.Text = "1 - Opera\xe7\xe3o Interna" Then cboNatureza.Text = "5102"\r\n'
    b'If cboDestOperacao.Text = "2 - Opera\xe7\xe3o Interestadual" Then cboNatureza.Text = "6102"\r\n'
    b'End Sub\r\n'
)
new_lostfocus_end = (
    b'If cboDestOperacao.Text = "1 - Opera\xe7\xe3o Interna" Then cboNatureza.Text = "5102"\r\n'
    b'If cboDestOperacao.Text = "2 - Opera\xe7\xe3o Interestadual" Then cboNatureza.Text = "6102"\r\n'
    b'AtualizarCFOPCSTItens\r\n'
    b'End Sub\r\n'
)
c = data.count(old_lostfocus_end)
print(f'2b. Chamar AtualizarCFOPCSTItens em LostFocus: {c}')
if c == 1: data = data.replace(old_lostfocus_end, new_lostfocus_end)

# Normalizar CRLF
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
