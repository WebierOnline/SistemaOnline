data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# Adicionar Dim para variaveis ST no bloco de Dims do Load_Data_Itens
old_dims = b'Dim curBaseICMS As Currency\r\n'
new_dims = (
    b'Dim curBaseICMS As Currency\r\n'
    b'Dim curVBCST As Currency\r\n'
    b'Dim curVICMSST As Currency\r\n'
    b'Dim dblMVAFinal As Double\r\n'
    b'Dim dblAliqInter As Double\r\n'
    b'Dim dblAliqInterna As Double\r\n'
)
c = data.count(old_dims)
print(f'1. Dims ST calc: {c}')
if c == 1: data = data.replace(old_dims, new_dims)

# Inserir secao 'ICMS-ST apos Tb("referencia")
old_ref = b'    Tb("referencia") = Format(0, "@")\r\n'
new_ref = (
    b'    Tb("referencia") = Format(0, "@")\r\n'
    b'\r\n'
    b"    'ICMS-ST\r\n"
    b'    If chkICMSST.Value = 1 Then\r\n'
    b"        ' chkICMSST marcado: calcula ou copia ST dependendo do regime e finalidade\r\n"
    b'        If (vRegimeTributario = 1 Or vRegimeTributario = 2 Or vRegimeTributario = 5) And Left(cboFinalidade.Text, 1) <> "4" Then\r\n'
    b"            ' Simples Nacional venda normal: zera ST (ST ja retido anteriormente)\r\n"
    b'            Tb("modBCST") = Format(0, "@")\r\n'
    b'            Tb("pMVAST") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("pRedBCST") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("vBCST") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("pICMSST") = CDbl(Format(0, "@"))\r\n'
    b'            Tb("vICMSST") = CDbl(Format(0, "@"))\r\n'
    b'        Else\r\n'
    b"            ' Regime Normal ou devolucao: calcula ST via MVA\r\n"
    b'            dblAliqInterna = CDbl(IIf(vPICMSST = "" Or vPICMSST = "0,00", 0, vPICMSST))\r\n'
    b'            dblAliqInter = CDbl(IIf(vAliqUFInter = 0, 0, vAliqUFInter))\r\n'
    b'\r\n'
    b"            ' MVA: ajustado se interestadual, original se interna\r\n"
    b'            If vUFEmpresa <> vUFDest And vUFDest <> "" And dblAliqInterna > 0 Then\r\n'
    b'                Dim dblMVAOrig As Double\r\n'
    b'                dblMVAOrig = CDbl(IIf(vPMVAST = "" Or vPMVAST = "0,00", 0, vPMVAST))\r\n'
    b'                If (1 - dblAliqInterna / 100) <> 0 Then\r\n'
    b'                    dblMVAFinal = (((1 + dblMVAOrig / 100) * (1 - dblAliqInter / 100)) / (1 - dblAliqInterna / 100) - 1) * 100\r\n'
    b'                    dblMVAFinal = Int(dblMVAFinal * 100 + 0.5) / 100\r\n'
    b'                Else\r\n'
    b'                    dblMVAFinal = dblMVAOrig\r\n'
    b'                End If\r\n'
    b'            Else\r\n'
    b'                dblMVAFinal = CDbl(IIf(vPMVAST = "" Or vPMVAST = "0,00", 0, vPMVAST))\r\n'
    b'            End If\r\n'
    b'\r\n'
    b"            ' Base do ST: (valor liquido + IPI) * (1 + MVA/100)\r\n"
    b'            curVBCST = (CCur(vValorProdutos) + vValorIPI) * (1 + dblMVAFinal / 100)\r\n'
    b'\r\n'
    b"            ' Reducao da base ST se houver\r\n"
    b'            If vPRedBCST <> "" And CDbl(vPRedBCST) > 0 Then\r\n'
    b'                curVBCST = curVBCST * (1 - CDbl(vPRedBCST) / 100)\r\n'
    b'            End If\r\n'
    b'\r\n'
    b"            ' vICMSST = (vBCST * aliq. interna) - ICMS proprio; nunca negativo\r\n"
    b'            curVICMSST = (curVBCST * dblAliqInterna / 100) - vValorICMS\r\n'
    b'            If curVICMSST < 0 Then curVICMSST = 0\r\n'
    b'\r\n'
    b'            Tb("modBCST") = Format(4, "@")\r\n'
    b'            Tb("pMVAST") = CDbl(Format(dblMVAFinal, "0.00"))\r\n'
    b'            Tb("pRedBCST") = CDbl(IIf(vPRedBCST = "", 0, Format(vPRedBCST, "@")))\r\n'
    b'            Tb("vBCST") = CDbl(Format(curVBCST, "0.00"))\r\n'
    b'            Tb("pICMSST") = CDbl(Format(dblAliqInterna, "0.00"))\r\n'
    b'            Tb("vICMSST") = CDbl(Format(curVICMSST, "0.00"))\r\n'
    b'        End If\r\n'
    b'    Else\r\n'
    b"        ' chkICMSST desmarcado: zera todos os campos ST\r\n"
    b'        Tb("modBCST") = Format(0, "@")\r\n'
    b'        Tb("pMVAST") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("pRedBCST") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("vBCST") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("pICMSST") = CDbl(Format(0, "@"))\r\n'
    b'        Tb("vICMSST") = CDbl(Format(0, "@"))\r\n'
    b'    End If\r\n'
)
c = data.count(old_ref)
print(f'2. Secao ICMS-ST: {c}')
if c == 1: data = data.replace(old_ref, new_ref)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
