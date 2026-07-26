path = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
with open(path, 'rb') as f:
    data = f.read()
content = data.decode('windows-1252')

R = '\r\n'

old = (
    '    curBCIS = curBCCBSIBS' + R +
    '    curVIS = CCur(Format(curBCIS * CDbl(IIf(vISpAliq = "", 0, vISpAliq)) / 100, "0.00"))' + R +
    '    ' + R +
    '    Tb("IBSCBS_CST") = Format(vIBSCBSCST, "@")' + R +
    '    Tb("IBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("IBS_pRed") = CDbl(IIf(vIBSpRed = "", 0, Format(vIBSpRed, "@")))' + R +
    '    Tb("IBS_UFpAliq") = CDbl(IIf(vIBSUFpAliq = "", 0, Format(vIBSUFpAliq, "@")))' + R +
    '    Tb("IBS_MunpAliq") = CDbl(IIf(vIBSMunpAliq = "", 0, Format(vIBSMunpAliq, "@")))' + R +
    '    Tb("IBS_vIBSUF") = CDbl(Format(curVIBSUF, "0.00"))' + R +
    '    Tb("IBS_vIBSMun") = CDbl(Format(curVIBSMun, "0.00"))' + R +
    '    Tb("IBS_vIBS") = CDbl(Format(curVIBSUF + curVIBSMun, "0.00"))' + R +
    '    Tb("CBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("CBS_pAliq") = CDbl(IIf(vCBSpAliq = "", 0, Format(vCBSpAliq, "@")))' + R +
    '    Tb("CBS_pRed") = CDbl(IIf(vCBSpRed = "", 0, Format(vCBSpRed, "@")))' + R +
    '    Tb("CBS_vCBS") = CDbl(Format(curVCBS, "0.00"))' + R +
    '    Tb("IS_CST") = Format(vISCST, "@")' + R +
    '    Tb("IS_tipo_calculo") = IIf(vISTipoCalc = "", 0, Val(vISTipoCalc))' + R +
    '    Tb("IS_vBC") = CDbl(Format(curBCIS, "0.00"))' + R +
    '    Tb("IS_pAliq") = CDbl(IIf(vISpAliq = "", 0, Format(vISpAliq, "@")))' + R +
    '    Tb("IS_qUnid") = CDbl(Format(0, "0.0000"))' + R +
    '    Tb("IS_vUnid") = CDbl(Format(0, "0.0000"))' + R +
    '    Tb("IS_vIS") = CDbl(Format(curVIS, "0.00"))' + R +
    'End Sub'
)

new = (
    '    ' + R +
    '    Tb("IBSCBS_CST") = Format(vIBSCBSCST, "@")' + R +
    '    Tb("IBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("IBS_pRed") = CDbl(IIf(vIBSpRed = "", 0, Format(vIBSpRed, "@")))' + R +
    '    Tb("IBS_UFpAliq") = CDbl(IIf(vIBSUFpAliq = "", 0, Format(vIBSUFpAliq, "@")))' + R +
    '    Tb("IBS_MunpAliq") = CDbl(IIf(vIBSMunpAliq = "", 0, Format(vIBSMunpAliq, "@")))' + R +
    '    Tb("IBS_vIBSUF") = CDbl(Format(curVIBSUF, "0.00"))' + R +
    '    Tb("IBS_vIBSMun") = CDbl(Format(curVIBSMun, "0.00"))' + R +
    '    Tb("IBS_vIBS") = CDbl(Format(curVIBSUF + curVIBSMun, "0.00"))' + R +
    '    Tb("CBS_vBC") = CDbl(Format(curBCCBSIBS, "0.00"))' + R +
    '    Tb("CBS_pAliq") = CDbl(IIf(vCBSpAliq = "", 0, Format(vCBSpAliq, "@")))' + R +
    '    Tb("CBS_pRed") = CDbl(IIf(vCBSpRed = "", 0, Format(vCBSpRed, "@")))' + R +
    '    Tb("CBS_vCBS") = CDbl(Format(curVCBS, "0.00"))' + R +
    '    ' + R +
    '    \' IS: calculo por tipo (0=zero, 1=ad valorem, 2=especifica, 3=mista)' + R +
    '    Dim iTipoIS As Integer' + R +
    '    iTipoIS = IIf(vISTipoCalc = "", 0, Val(vISTipoCalc))' + R +
    '    ' + R +
    '    Dim dQtdeIS As Double' + R +
    '    dQtdeIS = IIf(txtQuant.Text = "", 1, Val(Replace(Replace(txtQuant.Text, ".", ""), ",", ".")))' + R +
    '    If dQtdeIS = 0 Then dQtdeIS = 1' + R +
    '    ' + R +
    '    Dim dFatorConvIS As Double' + R +
    '    dFatorConvIS = Val(Replace(Replace(vISFatorConv, ".", ""), ",", "."))' + R +
    '    If dFatorConvIS = 0 Then dFatorConvIS = 1' + R +
    '    ' + R +
    '    Dim curISqUnid As Currency' + R +
    '    curISqUnid = CCur(Format(dQtdeIS * dFatorConvIS, "0.0000"))' + R +
    '    ' + R +
    '    Dim curISvUnid As Currency' + R +
    '    curISvUnid = CCur(IIf(vISvUnid = "", 0, Val(Replace(Replace(vISvUnid, ".", ""), ",", "."))))' + R +
    '    ' + R +
    '    Dim curISvBC As Currency' + R +
    '    If iTipoIS = 1 Or iTipoIS = 3 Then' + R +
    '        curISvBC = curBCCBSIBS' + R +
    '    Else' + R +
    '        curISvBC = 0' + R +
    '    End If' + R +
    '    ' + R +
    '    Dim dISpAliq As Double' + R +
    '    dISpAliq = IIf(vISpAliq = "", 0, Val(Replace(Replace(vISpAliq, ".", ""), ",", ".")))' + R +
    '    ' + R +
    '    Dim curISvIS As Currency' + R +
    '    Select Case iTipoIS' + R +
    '        Case 1' + R +
    '            curISvIS = CCur(Format(curISvBC * dISpAliq / 100, "0.00"))' + R +
    '        Case 2' + R +
    '            curISvIS = CCur(Format(curISqUnid * curISvUnid, "0.00"))' + R +
    '        Case 3' + R +
    '            curISvIS = CCur(Format(curISvBC * dISpAliq / 100 + curISqUnid * curISvUnid, "0.00"))' + R +
    '        Case Else' + R +
    '            curISvIS = 0' + R +
    '    End Select' + R +
    '    ' + R +
    '    Tb("IS_CST") = Format(vISCST, "@")' + R +
    '    Tb("IS_tipo_calculo") = iTipoIS' + R +
    '    Tb("IS_vBC") = CDbl(Format(curISvBC, "0.00"))' + R +
    '    Tb("IS_pAliq") = dISpAliq' + R +
    '    Tb("IS_qUnid") = CDbl(Format(curISqUnid, "0.0000"))' + R +
    '    Tb("IS_vUnid") = CDbl(Format(curISvUnid, "0.0000"))' + R +
    '    Tb("IS_vIS") = CDbl(Format(curISvIS, "0.00"))' + R +
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
