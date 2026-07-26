path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

old = (
    '    \' IBS/CBS/IS' + R +
    '    Dim curBCCBSIBS As Currency' + R +
    '    Dim curVIBSUF As Currency, curVIBSMun As Currency, curVCBS As Currency' + R +
    '    Dim curBCIS As Currency, curVIS As Currency' + R +
    '    ' + R +
    '    curBCCBSIBS = CCur(txtSubTotal.Text)' + R +
    '    curVIBSUF = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSUFpAliq = "", 0, vIBSUFpAliq)) / 100, "0.00"))' + R +
    '    curVIBSMun = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSMunpAliq = "", 0, vIBSMunpAliq)) / 100, "0.00"))' + R +
    '    curVCBS = CCur(Format(curBCCBSIBS * CDbl(IIf(vCBSpAliq = "", 0, vCBSpAliq)) / 100, "0.00"))' + R +
    '    curBCIS = curBCCBSIBS' + R +
    '    curVIS = CCur(Format(curBCIS * CDbl(IIf(vISpAliq = "", 0, vISpAliq)) / 100, "0.00"))' + R +
    '    ' + R +
    '    Tb("IBSCBS_CST") = Format(vIBSCBSCST, "@")' + R +
    '    Tb("CBS_pAliq") = CDbl(IIf(vCBSpAliq = "", 0, Format(vCBSpAliq, "@")))' + R +
    '    Tb("IBS_UFpAliq") = CDbl(IIf(vIBSUFpAliq = "", 0, Format(vIBSUFpAliq, "@")))' + R +
    '    Tb("IBS_MunpAliq") = CDbl(IIf(vIBSMunpAliq = "", 0, Format(vIBSMunpAliq, "@")))' + R +
    '    Tb("IBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("IBS_vIBSUF") = CDbl(Format(curVIBSUF, "0.00"))' + R +
    '    Tb("IBS_vIBSMun") = CDbl(Format(curVIBSMun, "0.00"))' + R +
    '    Tb("IBS_vIBS") = CDbl(Format(curVIBSUF + curVIBSMun, "0.00"))' + R +
    '    Tb("CBS_vCBS") = CDbl(Format(curVCBS, "0.00"))' + R +
    '    Tb("IS_CST") = Format(vISCST, "@")' + R +
    '    Tb("IS_pAliq") = CDbl(IIf(vISpAliq = "", 0, Format(vISpAliq, "@")))' + R +
    '    Tb("IS_vBC") = CDbl(Format(curBCIS, "0.00"))' + R +
    '    Tb("IS_vIS") = CDbl(Format(curVIS, "0.00"))' + R +
    'End Sub'
)

new = (
    '    \' IBS/CBS/IS' + R +
    '    Dim curBCCBSIBS As Currency' + R +
    '    Dim curVIBSUF As Currency, curVIBSMun As Currency, curVCBS As Currency' + R +
    '    Dim curBCIS As Currency, curVIS As Currency' + R +
    '    ' + R +
    '    curBCCBSIBS = CCur(txtSubTotal.Text) _' + R +
    '        + CCur(IIf(txtFrete.Text = "", 0, txtFrete.Text)) _' + R +
    '        + CCur(IIf(txtSeguro.Text = "", 0, txtSeguro.Text)) _' + R +
    '        + CCur(IIf(txtOutrosItem.Text = "", 0, txtOutrosItem.Text)) _' + R +
    '        - CCur(IIf(txtDesc.Text = "", 0, txtDesc.Text))' + R +
    '    curVIBSUF = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSUFpAliq = "", 0, vIBSUFpAliq)) / 100, "0.00"))' + R +
    '    curVIBSMun = CCur(Format(curBCCBSIBS * CDbl(IIf(vIBSMunpAliq = "", 0, vIBSMunpAliq)) / 100, "0.00"))' + R +
    '    curVCBS = CCur(Format(curBCCBSIBS * CDbl(IIf(vCBSpAliq = "", 0, vCBSpAliq)) / 100, "0.00"))' + R +
    '    curBCIS = curBCCBSIBS' + R +
    '    curVIS = CCur(Format(curBCIS * CDbl(IIf(vISpAliq = "", 0, vISpAliq)) / 100, "0.00"))' + R +
    '    ' + R +
    '    Tb("IBSCBS_CST") = Format(vIBSCBSCST, "@")' + R +
    '    Tb("IBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("IBS_pRed") = CDbl(Format(0, "0.00"))' + R +
    '    Tb("IBS_UFpAliq") = CDbl(IIf(vIBSUFpAliq = "", 0, Format(vIBSUFpAliq, "@")))' + R +
    '    Tb("IBS_MunpAliq") = CDbl(IIf(vIBSMunpAliq = "", 0, Format(vIBSMunpAliq, "@")))' + R +
    '    Tb("IBS_vIBSUF") = CDbl(Format(curVIBSUF, "0.00"))' + R +
    '    Tb("IBS_vIBSMun") = CDbl(Format(curVIBSMun, "0.00"))' + R +
    '    Tb("IBS_vIBS") = CDbl(Format(curVIBSUF + curVIBSMun, "0.00"))' + R +
    '    Tb("CBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("CBS_pAliq") = CDbl(IIf(vCBSpAliq = "", 0, Format(vCBSpAliq, "@")))' + R +
    '    Tb("CBS_pRed") = CDbl(Format(0, "0.00"))' + R +
    '    Tb("CBS_vCBS") = CDbl(Format(curVCBS, "0.00"))' + R +
    '    Tb("IS_CST") = Format(vISCST, "@")' + R +
    '    Tb("IS_tipo_calculo") = CDbl(Format(0, "@"))' + R +
    '    Tb("IS_vBC") = CDbl(Format(curBCIS, "0.00"))' + R +
    '    Tb("IS_pAliq") = CDbl(IIf(vISpAliq = "", 0, Format(vISpAliq, "@")))' + R +
    '    Tb("IS_qUnid") = CDbl(Format(0, "0.0000"))' + R +
    '    Tb("IS_vUnid") = CDbl(Format(0, "0.0000"))' + R +
    '    Tb("IS_vIS") = CDbl(Format(curVIS, "0.00"))' + R +
    'End Sub'
)

n = content.count(old)
if n != 1:
    print(f'ERRO: count={n}')
else:
    content = content.replace(old, new, 1)
    content = content.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
    with open(path, 'wb') as f:
        f.write(content.encode('windows-1252'))
    print('OK - arquivo gravado.')
