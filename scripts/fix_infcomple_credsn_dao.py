data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

old_sub = (
    b'Private Sub AtualizarInfCompleCredSN()\r\n'
    b'    If vTipoCRT = 3 Or txtCodNota.Text = "" Then Exit Sub\r\n'
    b'    Dim sBaseSimples As String\r\n'
    b'    sBaseSimples = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL N\xc3O GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."\r\n'
    b'    Dim rCred As DAO.Recordset\r\n'
    b'    Dim dblSomaCredSN As Double\r\n'
    b'    dblSomaCredSN = 0\r\n'
    b'    Set rCred = dbData.OpenRecordset("SELECT ISNULL(SUM(vCredICMSSN), 0) AS total FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND CST IN (\'101\',\'201\')")\r\n'
    b'    If Not rCred.EOF Then dblSomaCredSN = CDbl(rCred("total"))\r\n'
    b'    rCred.Close\r\n'
    b'    Set rCred = Nothing\r\n'
    b'    If dblSomaCredSN > 0 Then\r\n'
    b'        txtInfComple.Text = "DOCUMENTO EMITIDO POR ME OU EPP OPTANTE PELO SIMPLES NACIONAL. " & _\r\n'
    b'                            "PERMITE O APROVEITAMENTO DO CR\xc9DITO DE ICMS NO VALOR DE R$ " & Format(dblSomaCredSN, "#,##0.00") & _\r\n'
    b'                            "; CORRESPONDENTE \xc0 AL\xcdQUOTA DE " & Format(vPCreditoSN, "0.00") & _\r\n'
    b'                            "%, NOS TERMOS DO ART. 23 DA LEI COMPLEMENTAR N\xba 123, DE 2006."\r\n'
    b'    Else\r\n'
    b'        txtInfComple.Text = sBaseSimples\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
)

new_sub = (
    b'Private Sub AtualizarInfCompleCredSN()\r\n'
    b'    If vTipoCRT = 3 Or txtCodNota.Text = "" Then Exit Sub\r\n'
    b'    Dim sBaseSimples As String\r\n'
    b'    sBaseSimples = "EMPRESA ME OU EPP OPTANTE PELO SIMPLES NACIONAL N\xc3O GERA DIREITO A CREDITO FISCAL DE ICMS OU ISS."\r\n'
    b'    Dim rCred As ADODB.Recordset\r\n'
    b'    Dim dblSomaCredSN As Double\r\n'
    b'    dblSomaCredSN = 0\r\n'
    b'    Set rCred = dbData.OpenRecordset("SELECT ISNULL(SUM(vCredICMSSN), 0) AS total FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND CST IN (\'101\',\'201\')")\r\n'
    b'    If Not rCred.EOF Then dblSomaCredSN = CDbl(rCred("total"))\r\n'
    b'    rCred.Close\r\n'
    b'    Set rCred = Nothing\r\n'
    b'    If dblSomaCredSN > 0 Then\r\n'
    b'        txtInfComple.Text = "DOCUMENTO EMITIDO POR ME OU EPP OPTANTE PELO SIMPLES NACIONAL. " & _\r\n'
    b'                            "PERMITE O APROVEITAMENTO DO CR\xc9DITO DE ICMS NO VALOR DE R$ " & Format(dblSomaCredSN, "#,##0.00") & _\r\n'
    b'                            "; CORRESPONDENTE \xc0 AL\xcdQUOTA DE " & Format(vPCreditoSN, "0.00") & _\r\n'
    b'                            "%, NOS TERMOS DO ART. 23 DA LEI COMPLEMENTAR N\xba 123, DE 2006."\r\n'
    b'    Else\r\n'
    b'        txtInfComple.Text = sBaseSimples\r\n'
    b'    End If\r\n'
    b'End Sub\r\n'
)

c = data.count(old_sub)
print(f'AtualizarInfCompleCredSN: count={c}')
if c == 1:
    data = data.replace(old_sub, new_sub)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
