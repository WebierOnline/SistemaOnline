# -*- coding: utf-8 -*-
PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
data = open(PATH, 'rb').read()
text = data.decode('windows-1252')

patches = []

# ── Patch 1: cmdTransmitir_Click — verificar status antes de transmitir ───────
patches.append(('status_check',
    'If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub\r\n'
    '\r\n'
    "'verificar erros",

    'If vTipoEdicaoNFe = "Novo" Or vTipoEdicaoNFe = "Edicao" Then MsgBox "Existem um NFe em aberto, Salve-a ou Cancele-a!", vbExclamation, "Online Commerce": Frm_NF.Tab = 0: Exit Sub\r\n'
    '\r\n'
    'Dim vStatusNF As String\r\n'
    'vStatusNF = GridNotas.TextMatrix(GridNotas.Row, 7)\r\n'
    'If vStatusNF = "Enviada" Or vStatusNF = "Cancelada" Or vStatusNF = "Inutilizada" Then MsgBox "Esta nota j\xe1 foi " & vStatusNF & "!", vbExclamation, "Online Commerce": Exit Sub\r\n'
    '\r\n'
    "'verificar erros",
))

# ── Patch 2a: VerificarProdutosEnviar — corrigir condicao CFOP ────────────────
patches.append(('cfop_cond',
    "     'CFOP..........\r\n"
    '     If rNFCeItens!CFOP <> Empty Or rNFCeItens!CFOP = "0" Then\r\n'
    '         If Len(rNFCeItens!CFOP) > 4 Or Len(rNFCeItens!CFOP) < 4 Then\r\n'
    '             EncontroErroNFCe = True\r\n'
    '         Else\r\n'
    '             EncontroErroNFCe = False\r\n'
    '         End If\r\n'
    '     Else\r\n'
    '         EncontroErroNFCe = False\r\n'
    '     End If',

    "     'CFOP..........\r\n"
    '     EncontroErroNFCe = (Len(IIf(IsNull(rNFCeItens!CFOP), "", CStr(rNFCeItens!CFOP))) <> 4)',
))

# ── Patch 2b: VerificarProdutosEnviar — corrigir condicao NCM ────────────────
patches.append(('ncm_cond',
    "     'NCM..........\r\n"
    '     If rNFCeItens!NCM <> Empty Or rNFCeItens!NCM = "0" Or rNFCeItens!NCM = "" Then\r\n'
    '         If Len(rNFCeItens!NCM) > 8 Or Len(rNFCeItens!NCM) < 8 Then\r\n'
    '             EncontroErroNFCe = True\r\n'
    '         Else\r\n'
    '             EncontroErroNFCe = False\r\n'
    '         End If\r\n'
    '     Else\r\n'
    '         EncontroErroNFCe = False\r\n'
    '     End If\r\n'
    "     \r\n"
    "     If EncontroErroNFCe = True Then 'GoTo Continuar\r\n"
    '        If MsgBox("Produto com NCM incorreto ou inv\xe1lido! " & Chr(13) & " Produto: \'" & rNFCeItens!NomeProduto & "\' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then\r\n'
    '            vPossuiErro = True\r\n'
    '            GoTo Continuar\r\n'
    '        Else\r\n'
    '            vPossuiErro = True\r\n'
    '            Exit Sub\r\n'
    '        End If\r\n'
    '    End If\r\n'
    "     'End If\r\n"
    '     \r\n'
    "     'UNIDADE DE MEDIDA..........",

    "     'NCM..........\r\n"
    '     EncontroErroNFCe = (Len(IIf(IsNull(rNFCeItens!NCM), "", CStr(rNFCeItens!NCM))) <> 8)\r\n'
    '     \r\n'
    "     If EncontroErroNFCe = True Then 'GoTo Continuar\r\n"
    '        If MsgBox("Produto com NCM incorreto ou inv\xe1lido! " & Chr(13) & " Produto: \'" & rNFCeItens!NomeProduto & "\' " & Chr(13) & " Deseja corrigir esse produto?", vbQuestion + vbYesNo, "Erro") = vbYes Then\r\n'
    '            vPossuiErro = True\r\n'
    '            GoTo Continuar\r\n'
    '        Else\r\n'
    '            vPossuiErro = True\r\n'
    '            Exit Sub\r\n'
    '        End If\r\n'
    '    End If\r\n'
    '     \r\n'
    "     'UNIDADE DE MEDIDA..........",
))

# ── Patch 2c: VerificarProdutosEnviar — adicionar checks Quantidade e Valor ──
patches.append(('quant_valor_check',
    '        End If\r\n'
    '    End If\r\n'
    ' \r\n'
    ' rNFCeItens.MoveNext\r\n'
    ' Next\r\n'
    '    \r\n'
    'Continuar:',

    '        End If\r\n'
    '    End If\r\n'
    ' \r\n'
    "     'QUANTIDADE\r\n"
    '     If CDbl(IIf(IsNull(rNFCeItens!QuantidadeComercial), 0, rNFCeItens!QuantidadeComercial)) <= 0 Then\r\n'
    '         MsgBox "Produto \'" & rNFCeItens!NomeProduto & "\' com quantidade zero ou negativa!", vbExclamation, "Erro"\r\n'
    '         vPossuiErro = True\r\n'
    '         GoTo Continuar\r\n'
    '     End If\r\n'
    '     \r\n'
    "     'VALOR UNIT\xc1RIO\r\n"
    '     If CDbl(IIf(IsNull(rNFCeItens!ValorUnitarioComercializacao), 0, rNFCeItens!ValorUnitarioComercializacao)) <= 0 Then\r\n'
    '         MsgBox "Produto \'" & rNFCeItens!NomeProduto & "\' com valor unit\xe1rio zero ou negativo!", vbExclamation, "Erro"\r\n'
    '         vPossuiErro = True\r\n'
    '         GoTo Continuar\r\n'
    '     End If\r\n'
    ' \r\n'
    ' rNFCeItens.MoveNext\r\n'
    ' Next\r\n'
    '    \r\n'
    'Continuar:',
))

# ── Patch 3a: CorrecoesBasicasNFe — detectar FORNECEDOR ─────────────────────
patches.append(('fornecedor_detect',
    'IdNFProd = GridNotas.TextMatrix(GridNotas.Row, 1)\r\n'
    '\r\n'
    "'consultar o cpf/cnpj na nota",

    'IdNFProd = GridNotas.TextMatrix(GridNotas.Row, 1)\r\n'
    'Dim vTipoDestCorrec As String\r\n'
    'vTipoDestCorrec = GridNotas.TextMatrix(GridNotas.Row, 13)\r\n'
    '\r\n'
    "'consultar o cpf/cnpj na nota",
))

# ── Patch 3b: CorrecoesBasicasNFe — nao sair para FORNECEDOR ────────────────
patches.append(('fornecedor_exit',
    'If vCPF = "" Then Exit Sub',
    'If vCPF = "" And vTipoDestCorrec <> "FORNECEDOR" Then Exit Sub',
))

# ── Patch 4: CorrecoesBasicasNFe — expandir mapeamento CFOP (+ 5101/6101) ───
patches.append(('cfop_map_interstate_sn',
    '                If rNotaFiscalItens!CFOP = "5102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST = '500' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        Else                        'empresa lucro presumido ou lucro real",

    '                If rNotaFiscalItens!CFOP = "5101" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6101', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST = '500' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        Else                        'empresa lucro presumido ou lucro real",
))

patches.append(('cfop_map_interstate_normal',
    '                If rNotaFiscalItens!CFOP = "5102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        End If                      'empresa simples nacional FIM\r\n"
    '\r\n'
    '     rNotaFiscalItens.MoveNext\r\n'
    '     Next\r\n'
    'Else',

    '                If rNotaFiscalItens!CFOP = "5101" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6101' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5403" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '6403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        End If                      'empresa simples nacional FIM\r\n"
    '\r\n'
    '     rNotaFiscalItens.MoveNext\r\n'
    '     Next\r\n'
    'Else',
))

patches.append(('cfop_map_intrastate_sn',
    '                If rNotaFiscalItens!CFOP = "6102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "6405" Or rNotaFiscalItens!CFOP = "6403" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST = '500' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        Else                        'empresa lucro presumido ou lucro real",

    '                If rNotaFiscalItens!CFOP = "6101" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5101', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "6102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "6405" Or rNotaFiscalItens!CFOP = "6403" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST = '500' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102', CST = '102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        Else                        'empresa lucro presumido ou lucro real",
))

patches.append(('cfop_map_intrastate_normal',
    '                If rNotaFiscalItens!CFOP = "6102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Or rNotaFiscalItens!CFOP = "6403" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        End If                      'empresa simples nacional FIM",

    '                If rNotaFiscalItens!CFOP = "6101" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5101' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "6102" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5102' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "6403" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                ElseIf rNotaFiscalItens!CFOP = "5405" Or rNotaFiscalItens!CFOP = "6405" Then\r\n'
    "                    dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    '                End If\r\n'
    '            Else\r\n'
    "                dbData.Execute \"UPDATE NotaFiscalItens SET CFOP = '5403', CST= '060', pICMS = '0.00' WHERE CodigoProduto  = \" & vCodProdConsulta & \" And CodigoNota = \" & IdNFProd\r\n"
    "            End If\r\n"
    "        End If                      'empresa simples nacional FIM",
))

# ── Verificar e aplicar ────────────────────────────────────────────────────────
all_ok = True
for name, old, new in patches:
    c = text.count(old)
    status = 'OK' if c == 1 else f'ERRO count={c}'
    print(f'{name}: {status}')
    if c != 1:
        all_ok = False

if all_ok:
    for _, old, new in patches:
        text = text.replace(old, new)
    out = text.encode('windows-1252')
    out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
    open(PATH, 'wb').write(out)
    print('\nOK — todos os patches aplicados')
else:
    print('\nABORTADO')
