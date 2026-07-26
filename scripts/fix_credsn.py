data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# =============================================================================
# 1. Declaracao form-level: Dim vPCreditoSN As Double
# =============================================================================
old_dim = (
    b'Dim vIPICompoeDIFAL As Integer\r\n'
)
new_dim = (
    b'Dim vIPICompoeDIFAL As Integer\r\n'
    b'Dim vPCreditoSN     As Double\r\n'
)
c = data.count(old_dim)
print(f'1. Dim vPCreditoSN: {c}')
if c == 1: data = data.replace(old_dim, new_dim)

# =============================================================================
# 2. Empresa SELECT: incluir pCreditoICMSSimplesNacional (2 ocorrencias iguais)
# =============================================================================
old_emp_sel = (
    b'SELECT CRT, ESTADO, RegimeTributario, IPICompoeDIFAL FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
    b'    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))\r\n'
    b'    vIPICompoeDIFAL = IIf(IsNull(r("IPICompoeDIFAL")), 0, r("IPICompoeDIFAL"))\r\n'
)
new_emp_sel = (
    b'SELECT CRT, ESTADO, RegimeTributario, IPICompoeDIFAL, pCreditoICMSSimplesNacional FROM empresa"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'\r\n'
    b'If Not r.EOF Then\r\n'
    b'    vTipoCRT = r("CRT")\r\n'
    b'    vUFEmpresa = r("ESTADO")\r\n'
    b'    vRegimeTributario = IIf(IsNull(r("RegimeTributario")), 0, r("RegimeTributario"))\r\n'
    b'    vIPICompoeDIFAL = IIf(IsNull(r("IPICompoeDIFAL")), 0, r("IPICompoeDIFAL"))\r\n'
    b'    vPCreditoSN = IIf(IsNull(r("pCreditoICMSSimplesNacional")), 0, CDbl(r("pCreditoICMSSimplesNacional")))\r\n'
)
c = data.count(old_emp_sel)
print(f'2. Empresa SELECT pCreditoSN: {c}')
if c == 2: data = data.replace(old_emp_sel, new_emp_sel)

# =============================================================================
# 3. Load_Data_Itens: calcular pCredSN / vCredICMSSN antes do End Sub
# =============================================================================
old_ldi_end = (
    b'        Tb("vFCPUFDest") = 0: Tb("vICMSUFRemet") = 0: Tb("vICMSUFDest") = 0\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub Calcular_Total()\r\n'
)
new_ldi_end = (
    b'        Tb("vFCPUFDest") = 0: Tb("vICMSUFRemet") = 0: Tb("vICMSUFDest") = 0\r\n'
    b'    End If\r\n'
    b'    \r\n'
    b"    ' Credito Simples Nacional (CSOSN 101/201)\r\n"
    b'    If (vRegimeTributario = 1 Or vRegimeTributario = 2) And (sCSTFinal = "101" Or sCSTFinal = "201") Then\r\n'
    b'        Tb("pCredSN")     = vPCreditoSN\r\n'
    b'        Tb("vCredICMSSN") = CCur(Format(curBaseICMS * vPCreditoSN / 100, "0.00"))\r\n'
    b'    Else\r\n'
    b'        Tb("pCredSN")     = 0\r\n'
    b'        Tb("vCredICMSSN") = 0\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub Calcular_Total()\r\n'
)
c = data.count(old_ldi_end)
print(f'3. Load_Data_Itens pCredSN: {c}')
if c == 1: data = data.replace(old_ldi_end, new_ldi_end)

# =============================================================================
# 4. RecalcularItensNota: adicionar Dims + calculo + campos no UPDATE
# =============================================================================
# 4a. Adicionar Dim curPCredSN / curVCredICMSSN na lista de Dims
old_dim_rec = (
    b'Dim curVBCST     As Currency\r\n'
    b'Dim curVICMSST   As Currency\r\n'
)
new_dim_rec = (
    b'Dim curVBCST     As Currency\r\n'
    b'Dim curVICMSST   As Currency\r\n'
    b'Dim curPCredSN   As Double\r\n'
    b'Dim curVCredICMSSN As Currency\r\n'
)
c = data.count(old_dim_rec)
print(f'4a. Dims RecalcularItensNota: {c}')
if c == 1: data = data.replace(old_dim_rec, new_dim_rec)

# 4b. Inserir calculo apos bloco ICMS-ST (antes do sUpd)
old_before_upd = (
    b'    Else\r\n'
    b'        curVBCST   = 0\r\n'
    b'        curVICMSST = 0\r\n'
    b'    End If\r\n'
    b'\r\n'
    b'    sUpd = "UPDATE NotaFiscalItens SET " & _\r\n'
)
new_before_upd = (
    b'    Else\r\n'
    b'        curVBCST   = 0\r\n'
    b'        curVICMSST = 0\r\n'
    b'    End If\r\n'
    b'\r\n'
    b"    ' Credito Simples Nacional (CSOSN 101/201)\r\n"
    b'    If bSimples And Not bDevolucao And (sCST = "101" Or sCST = "201") Then\r\n'
    b'        curPCredSN     = vPCreditoSN\r\n'
    b'        curVCredICMSSN = CCur(Format(curBaseICMS * curPCredSN / 100, "0.00"))\r\n'
    b'    Else\r\n'
    b'        curPCredSN     = 0\r\n'
    b'        curVCredICMSSN = 0\r\n'
    b'    End If\r\n'
    b'\r\n'
    b'    sUpd = "UPDATE NotaFiscalItens SET " & _\r\n'
)
c = data.count(old_before_upd)
print(f'4b. Calculo CredSN em RecalcularItensNota: {c}')
if c == 1: data = data.replace(old_before_upd, new_before_upd)

# 4c. Adicionar pCredSN e vCredICMSSN no UPDATE
old_upd_end = (
    b'           "vBCST = " & FSQL(curVBCST, 2) & ", " & _\r\n'
    b'           "vICMSST = " & FSQL(curVICMSST, 2) & " " & _\r\n'
    b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & vItem\r\n'
)
new_upd_end = (
    b'           "vBCST = " & FSQL(curVBCST, 2) & ", " & _\r\n'
    b'           "vICMSST = " & FSQL(curVICMSST, 2) & ", " & _\r\n'
    b'           "pCredSN = " & FSQL(curPCredSN, 4) & ", " & _\r\n'
    b'           "vCredICMSSN = " & FSQL(curVCredICMSSN, 2) & " " & _\r\n'
    b'           "WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & vItem\r\n'
)
c = data.count(old_upd_end)
print(f'4c. UPDATE pCredSN/vCredICMSSN: {c}')
if c == 1: data = data.replace(old_upd_end, new_upd_end)

# =============================================================================
# 5. Novo sub AtualizarInfCompleCredSN + chamada em AtualizarTotaisNota
# =============================================================================
old_atu_totais_end = (
    b'If IsDate(mskEmissao) Then mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy") Else mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy")\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub Calcul'
)
new_atu_totais_end = (
    b'If IsDate(mskEmissao) Then mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy") Else mskInicioDup.Text = Format(mskEmissao.Text, "dd/mm/yy")\r\n'
    b'AtualizarInfCompleCredSN\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub AtualizarInfCompleCredSN()\r\n'
    b'    If vTipoCRT = 3 Or txtCodNota.Text = "" Then Exit Sub\r\n'
    b'    Dim sBaseSimples As String\r\n'
    b'    sBaseSimples = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL N\xc3O GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."\r\n'
    b'    Dim dblSomaCredSN As Double\r\n'
    b'    dblSomaCredSN = CDbl(SQLExecutaRetorno("SELECT ISNULL(SUM(vCredICMSSN), 0) r FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND CST IN (\'101\',\'201\')", "r", 0))\r\n'
    b'    If dblSomaCredSN > 0 Then\r\n'
    b'        txtInfComple.Text = "PERMITE O APROVEITAMENTO DO CR\xc9DITO DE ICMS NO VALOR DE R$ " & Format(dblSomaCredSN, "#,##0.00") & "; CORRESPONDENTE \xc0 AL\xcdQUOTA DE " & Format(vPCreditoSN, "0.00") & "%, NOS TERMOS DO ART. 23 DA LEI COMPLEMENTAR N\xba 123, DE 2006."\r\n'
    b'    Else\r\n'
    b'        txtInfComple.Text = sBaseSimples\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub Calcul'
)
c = data.count(old_atu_totais_end)
print(f'5. AtualizarInfCompleCredSN: {c}')
if c == 1: data = data.replace(old_atu_totais_end, new_atu_totais_end)

# Normalizar CRLF
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
