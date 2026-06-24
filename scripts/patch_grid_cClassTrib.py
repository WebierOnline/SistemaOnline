# -*- coding: utf-8 -*-
# Insere coluna 10 (cClassTrib) no GridNotasItens de NFe_Completa.frm
# Todas as colunas 10-30 (antigas) viram 11-31 (novas).
# Total de colunas: 31 -> 32.

import sys

PATH = r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm'
text = open(PATH, 'rb').read().decode('windows-1252')
ok = True

def crlf(s):
    return s.replace('\n', '\r\n')

patches = []

# ══════════════════════════════════════════════════════════════
# A1 – .Cols = 31 → 32
# ══════════════════════════════════════════════════════════════
patches.append(('A1 .Cols=32', 1,
    crlf('      .Clear\n      .Cols = 31\n      .rows = 2'),
    crlf('      .Clear\n      .Cols = 32\n      .rows = 2')
))

# ══════════════════════════════════════════════════════════════
# A2 – ColWidth block (reforma + resto)
# ══════════════════════════════════════════════════════════════
patches.append(('A2 ColWidth block', 1,
crlf("""\
      'Reforma tributaria (chkReforma) - ocultas por padrao
      .ColWidth(9) = 0      'CST IBS/CBS
      .ColWidth(10) = 0     'V. IBS
      .ColWidth(11) = 0     'V. CBS
      .ColWidth(12) = 0     'V. IS
      .ColWidth(13) = 850   'VALOR
      .ColWidth(14) = 850   'QTDE
      .ColWidth(15) = 800   'FRETE
      .ColWidth(16) = 0     'SEGURO (chkSeguro) - oculto por padrao
      .ColWidth(17) = 0     'OUTROS (chkOutros) - oculto por padrao
      .ColWidth(18) = 800   'DESC.
      .ColWidth(19) = 1050  'TOTAL
      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)
      .ColWidth(20) = 0     'BC ICMS
      .ColWidth(21) = 0     '%ICMS
      .ColWidth(22) = 0     'ICMS
      .ColWidth(23) = 0     '%RED BC
      .ColWidth(24) = 0     'BC ST
      .ColWidth(25) = 0     '%ICMSST
      .ColWidth(26) = 0     'ICMSST
      .ColWidth(27) = 0     'MVA ST
      .ColWidth(28) = 0     'CST IPI
      .ColWidth(29) = 0     '%IPI
      .ColWidth(30) = 0     'cEnq"""),
crlf("""\
      'Reforma tributaria (chkReforma) - ocultas por padrao
      .ColWidth(9) = 0      'CST IBS/CBS
      .ColWidth(10) = 0     'cClassTrib (chkReforma)
      .ColWidth(11) = 0     'V. IBS
      .ColWidth(12) = 0     'V. CBS
      .ColWidth(13) = 0     'V. IS
      .ColWidth(14) = 850   'VALOR
      .ColWidth(15) = 850   'QTDE
      .ColWidth(16) = 800   'FRETE
      .ColWidth(17) = 0     'SEGURO (chkSeguro) - oculto por padrao
      .ColWidth(18) = 0     'OUTROS (chkOutros) - oculto por padrao
      .ColWidth(19) = 800   'DESC.
      .ColWidth(20) = 1050  'TOTAL
      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)
      .ColWidth(21) = 0     'BC ICMS
      .ColWidth(22) = 0     '%ICMS
      .ColWidth(23) = 0     'ICMS
      .ColWidth(24) = 0     '%RED BC
      .ColWidth(25) = 0     'BC ST
      .ColWidth(26) = 0     '%ICMSST
      .ColWidth(27) = 0     'ICMSST
      .ColWidth(28) = 0     'MVA ST
      .ColWidth(29) = 0     'CST IPI
      .ColWidth(30) = 0     '%IPI
      .ColWidth(31) = 0     'IPI""")
))

# ══════════════════════════════════════════════════════════════
# A3 – TextMatrix headers (cols 9-30)
# ══════════════════════════════════════════════════════════════
patches.append(('A3 TextMatrix headers', 1,
crlf("""\
      .TextMatrix(0, 9) = "CST IBS"
      .TextMatrix(0, 10) = "V. IBS"
      .TextMatrix(0, 11) = "V. CBS"
      .TextMatrix(0, 12) = "V. IS"
      .TextMatrix(0, 13) = "VALOR"
      .TextMatrix(0, 14) = "QTDE"
      .TextMatrix(0, 15) = "FRETE"
      .TextMatrix(0, 16) = "SEGURO"
      .TextMatrix(0, 17) = "OUTROS"
      .TextMatrix(0, 18) = "DESC."
      .TextMatrix(0, 19) = "TOTAL"
      .TextMatrix(0, 20) = "BC ICMS"
      .TextMatrix(0, 21) = "%ICMS"
      .TextMatrix(0, 22) = "ICMS"
      .TextMatrix(0, 23) = "%RED BC"
      .TextMatrix(0, 24) = "BC ST"
      .TextMatrix(0, 25) = "%ICMSST"
      .TextMatrix(0, 26) = "ICMSST"
      .TextMatrix(0, 27) = "MVA ST"
      .TextMatrix(0, 28) = "CST IPI"
      .TextMatrix(0, 29) = "%IPI"
      .TextMatrix(0, 30) = "IPI\""""),
crlf("""\
      .TextMatrix(0, 9) = "CST IBS"
      .TextMatrix(0, 10) = "CLASS. IBS"
      .TextMatrix(0, 11) = "V. IBS"
      .TextMatrix(0, 12) = "V. CBS"
      .TextMatrix(0, 13) = "V. IS"
      .TextMatrix(0, 14) = "VALOR"
      .TextMatrix(0, 15) = "QTDE"
      .TextMatrix(0, 16) = "FRETE"
      .TextMatrix(0, 17) = "SEGURO"
      .TextMatrix(0, 18) = "OUTROS"
      .TextMatrix(0, 19) = "DESC."
      .TextMatrix(0, 20) = "TOTAL"
      .TextMatrix(0, 21) = "BC ICMS"
      .TextMatrix(0, 22) = "%ICMS"
      .TextMatrix(0, 23) = "ICMS"
      .TextMatrix(0, 24) = "%RED BC"
      .TextMatrix(0, 25) = "BC ST"
      .TextMatrix(0, 26) = "%ICMSST"
      .TextMatrix(0, 27) = "ICMSST"
      .TextMatrix(0, 28) = "MVA ST"
      .TextMatrix(0, 29) = "CST IPI"
      .TextMatrix(0, 30) = "%IPI"
      .TextMatrix(0, 31) = "IPI\"""")
))

# ══════════════════════════════════════════════════════════════
# A4 – alinhamento loop 9 To 30 → 9 To 31
# ══════════════════════════════════════════════════════════════
patches.append(('A4 alignment loop', 1,
    crlf("      'Alinhamento: texto esquerda (0-8), numeros direita (9-30)\n      For i = 0 To 8\n         .ColAlignment(i) = 1\n      Next i\n      For i = 9 To 30\n         .ColAlignment(i) = 6\n      Next i"),
    crlf("      'Alinhamento: texto esquerda (0-8), numeros direita (9-31)\n      For i = 0 To 8\n         .ColAlignment(i) = 1\n      Next i\n      For i = 9 To 31\n         .ColAlignment(i) = 6\n      Next i")
))

# ══════════════════════════════════════════════════════════════
# A5 – data fill loop (cols 9-30)
# ══════════════════════════════════════════════════════════════
patches.append(('A5 data fill loop', 1,
crlf("""\
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("IBSCBS_CST"))
            .TextMatrix(.rows - 1, 10) = FormatNumber(rTabela("IBS_vIBS"), 2)
            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("CBS_vCBS"), 2)
            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("IS_vIS"), 2)
            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela("ValorUnitarioComercializacao"), 2)
            If rTabela("UnidadeComercial") = "KG" Or rTabela("UnidadeComercial") = "GR" Or rTabela("UnidadeComercial") = "MG" Then
                .TextMatrix(.rows - 1, 14) = Format(rTabela("QuantidadeComercial"), ocPESO)
            Else
                .TextMatrix(.rows - 1, 14) = Format(rTabela("QuantidadeComercial"), "###,###,##0")
            End If
            .TextMatrix(.rows - 1, 15) = FormatNumber(rTabela("ValorFrete"), 2)
            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela("ValorSeguro"), 2)
            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela("ValorOutros"), 2)
            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela("ValorDesconto"), 2)
            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela("ValorTotalBruto"), 2)
            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela("vBC"), 2)
            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela("pICMS"), 2)
            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela("vICMS"), 2)
            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela("pRedBC"), 2)
            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela("vBCST"), 2)
            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("pICMSST"), 2)
            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela("vICMSST"), 2)
            .TextMatrix(.rows - 1, 27) = FormatNumber(rTabela("pMVAST"), 2)
            .TextMatrix(.rows - 1, 28) = rTabela("IPICST")
            .TextMatrix(.rows - 1, 29) = FormatNumber(rTabela("IPIpIPI"), 2)
            .TextMatrix(.rows - 1, 30) = FormatNumber(rTabela("IPIvIPI"), 2)"""),
crlf("""\
            .TextMatrix(.rows - 1, 9) = ValidateNull(rTabela("IBSCBS_CST"))
            .TextMatrix(.rows - 1, 10) = ValidateNull(rTabela("cClassTrib"))
            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("IBS_vIBS"), 2)
            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("CBS_vCBS"), 2)
            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela("IS_vIS"), 2)
            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("ValorUnitarioComercializacao"), 2)
            If rTabela("UnidadeComercial") = "KG" Or rTabela("UnidadeComercial") = "GR" Or rTabela("UnidadeComercial") = "MG" Then
                .TextMatrix(.rows - 1, 15) = Format(rTabela("QuantidadeComercial"), ocPESO)
            Else
                .TextMatrix(.rows - 1, 15) = Format(rTabela("QuantidadeComercial"), "###,###,##0")
            End If
            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela("ValorFrete"), 2)
            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela("ValorSeguro"), 2)
            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela("ValorOutros"), 2)
            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela("ValorDesconto"), 2)
            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela("ValorTotalBruto"), 2)
            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela("vBC"), 2)
            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela("pICMS"), 2)
            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela("vICMS"), 2)
            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela("pRedBC"), 2)
            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("vBCST"), 2)
            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela("pICMSST"), 2)
            .TextMatrix(.rows - 1, 27) = FormatNumber(rTabela("vICMSST"), 2)
            .TextMatrix(.rows - 1, 28) = FormatNumber(rTabela("pMVAST"), 2)
            .TextMatrix(.rows - 1, 29) = rTabela("IPICST")
            .TextMatrix(.rows - 1, 30) = FormatNumber(rTabela("IPIpIPI"), 2)
            .TextMatrix(.rows - 1, 31) = FormatNumber(rTabela("IPIvIPI"), 2)""")
))

# ══════════════════════════════════════════════════════════════
# A6 – cor editable Array (shift + add col 10)
# ══════════════════════════════════════════════════════════════
patches.append(('A6 colEdit Array', 1,
    'For Each colEdit In Array(2, 5, 6, 7, 8, 21, 23, 25, 27, 28, 29)',
    'For Each colEdit In Array(2, 5, 6, 7, 8, 10, 22, 24, 26, 28, 29, 30)'
))

# ══════════════════════════════════════════════════════════════
# A7 – cor reforma Array (shift cols 10-12 → 11-13)
# ══════════════════════════════════════════════════════════════
patches.append(('A7 colRef Array', 1,
    'For Each colRef In Array(9, 10, 11, 12)',
    'For Each colRef In Array(9, 11, 12, 13)'
))

# ══════════════════════════════════════════════════════════════
# B1 – Exibir_Itens SELECT adiciona cClassTrib
# ══════════════════════════════════════════════════════════════
patches.append(('B1 SELECT cClassTrib', 1,
    '"IBSCBS_CST, IBS_vIBS, CBS_vCBS, IS_vIS, " & _',
    '"IBSCBS_CST, cClassTrib, IBS_vIBS, CBS_vCBS, IS_vIS, " & _'
))

# ══════════════════════════════════════════════════════════════
# C – AplicarVisibilidadeGridItens (bloco inteiro)
# ══════════════════════════════════════════════════════════════
patches.append(('C AplicarVisibilidade', 1,
crlf("""\
Sub AplicarVisibilidadeGridItens()
   If GridNotasItens.Cols < 31 Then Exit Sub
   'Reforma tributaria: chkReforma
   Dim bReforma As Boolean
   bReforma = (chkReforma.Value = 1)
   GridNotasItens.ColWidth(9) = IIf(bReforma, 700, 0)    'CST IBS/CBS
   GridNotasItens.ColWidth(10) = IIf(bReforma, 850, 0)   'V. IBS
   GridNotasItens.ColWidth(11) = IIf(bReforma, 850, 0)   'V. CBS
   GridNotasItens.ColWidth(12) = IIf(bReforma, 850, 0)   'V. IS

   'Seguro: chkSeguro
   GridNotasItens.ColWidth(16) = IIf(chkSeguro.Value = 1, 900, 0)
   'Outros: chkOutros
   GridNotasItens.ColWidth(17) = IIf(chkOutros.Value = 1, 900, 0)

   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)
   Dim bICMS As Boolean
   bICMS = (Left(cboFinalidade.Text, 1) = "4")
   GridNotasItens.ColWidth(20) = IIf(bICMS, 850, 0)  'BC ICMS
   GridNotasItens.ColWidth(21) = IIf(bICMS, 850, 0)  '%ICMS
   GridNotasItens.ColWidth(22) = IIf(bICMS, 850, 0)  'ICMS

   '%RedBC: chkpRedBC
   GridNotasItens.ColWidth(23) = IIf(chkpRedBC.Value = 1, 700, 0)

   'Grupo ICMSST: chkICMSST
   Dim bST As Boolean
   bST = (chkICMSST.Value = 1)
   GridNotasItens.ColWidth(24) = IIf(bST, 850, 0)  'BC ST
   GridNotasItens.ColWidth(25) = IIf(bST, 900, 0)  '%ICMSST
   GridNotasItens.ColWidth(26) = IIf(bST, 850, 0)  'ICMSST
   GridNotasItens.ColWidth(27) = IIf(bST, 850, 0)  'MVA ST

   'Grupo IPI: chkIPI
   Dim bIPI As Boolean
   bIPI = (chkIPI.Value = 1)
   GridNotasItens.ColWidth(28) = IIf(bIPI, 850, 0)  'CST IPI
   GridNotasItens.ColWidth(29) = IIf(bIPI, 850, 0)  '%IPI
   GridNotasItens.ColWidth(30) = IIf(bIPI, 850, 0)  'cEnq
End Sub"""),
crlf("""\
Sub AplicarVisibilidadeGridItens()
   If GridNotasItens.Cols < 32 Then Exit Sub
   'Reforma tributaria: chkReforma
   Dim bReforma As Boolean
   bReforma = (chkReforma.Value = 1)
   GridNotasItens.ColWidth(9) = IIf(bReforma, 700, 0)     'CST IBS/CBS
   GridNotasItens.ColWidth(10) = IIf(bReforma, 1200, 0)   'cClassTrib
   GridNotasItens.ColWidth(11) = IIf(bReforma, 850, 0)    'V. IBS
   GridNotasItens.ColWidth(12) = IIf(bReforma, 850, 0)    'V. CBS
   GridNotasItens.ColWidth(13) = IIf(bReforma, 850, 0)    'V. IS

   'Seguro: chkSeguro
   GridNotasItens.ColWidth(17) = IIf(chkSeguro.Value = 1, 900, 0)
   'Outros: chkOutros
   GridNotasItens.ColWidth(18) = IIf(chkOutros.Value = 1, 900, 0)

   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)
   Dim bICMS As Boolean
   bICMS = (Left(cboFinalidade.Text, 1) = "4")
   GridNotasItens.ColWidth(21) = IIf(bICMS, 850, 0)  'BC ICMS
   GridNotasItens.ColWidth(22) = IIf(bICMS, 850, 0)  '%ICMS
   GridNotasItens.ColWidth(23) = IIf(bICMS, 850, 0)  'ICMS

   '%RedBC: chkpRedBC
   GridNotasItens.ColWidth(24) = IIf(chkpRedBC.Value = 1, 700, 0)

   'Grupo ICMSST: chkICMSST
   Dim bST As Boolean
   bST = (chkICMSST.Value = 1)
   GridNotasItens.ColWidth(25) = IIf(bST, 850, 0)  'BC ST
   GridNotasItens.ColWidth(26) = IIf(bST, 900, 0)  '%ICMSST
   GridNotasItens.ColWidth(27) = IIf(bST, 850, 0)  'ICMSST
   GridNotasItens.ColWidth(28) = IIf(bST, 850, 0)  'MVA ST

   'Grupo IPI: chkIPI
   Dim bIPI As Boolean
   bIPI = (chkIPI.Value = 1)
   GridNotasItens.ColWidth(29) = IIf(bIPI, 850, 0)  'CST IPI
   GridNotasItens.ColWidth(30) = IIf(bIPI, 850, 0)  '%IPI
   GridNotasItens.ColWidth(31) = IIf(bIPI, 850, 0)  'IPI
End Sub""")
))

# ══════════════════════════════════════════════════════════════
# D – GridNotasItens_Click: Case list
# ══════════════════════════════════════════════════════════════
patches.append(('D Click Case list', 1,
    '    Case 2, 5, 6, 7, 8, 21, 23, 25, 27, 28, 29',
    '    Case 2, 5, 6, 7, 8, 10, 22, 24, 26, 28, 29, 30'
))

# ══════════════════════════════════════════════════════════════
# E1 – Case 5 UND: col 14 (QTDE) → 15
# ══════════════════════════════════════════════════════════════
patches.append(('E1 Case5 QTDE col14->15', 1,
crlf("""\
        Dim sQtdAtual As String
        sQtdAtual = GridNotasItens.TextMatrix(iRow, 14)
        If sVal = "KG" Or sVal = "GR" Or sVal = "MG" Then
            GridNotasItens.TextMatrix(iRow, 14) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), ocPESO)
        Else
            GridNotasItens.TextMatrix(iRow, 14) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), "###,###,##0")
        End If"""),
crlf("""\
        Dim sQtdAtual As String
        sQtdAtual = GridNotasItens.TextMatrix(iRow, 15)
        If sVal = "KG" Or sVal = "GR" Or sVal = "MG" Then
            GridNotasItens.TextMatrix(iRow, 15) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), ocPESO)
        Else
            GridNotasItens.TextMatrix(iRow, 15) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), "###,###,##0")
        End If""")
))

# ══════════════════════════════════════════════════════════════
# E2 – Case 10 cClassTrib (inserir após Case 8)
# ══════════════════════════════════════════════════════════════
patches.append(('E2 insert Case10', 1,
crlf("""\
    Case 8 ' CST
        If sVal = "" Or Len(sVal) <> 3 Then
            MsgBox "CST deve ter 3 d\xedgitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET CST = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 21 ' %ICMS"""),
crlf("""\
    Case 8 ' CST
        If sVal = "" Or Len(sVal) <> 3 Then
            MsgBox "CST deve ter 3 d\xedgitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET CST = '" & sVal & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 10 ' cClassTrib
        sVal = UCase(Trim(sVal))
        dbData.Execute "UPDATE NotaFiscalItens SET cClassTrib = '" & Replace(sVal, "'", "''") & "' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal

    Case 22 ' %ICMS""")
))

# ══════════════════════════════════════════════════════════════
# E3 – Case 21 → 22 (interno: cols 20→21, 21→22, 22→23)
# ══════════════════════════════════════════════════════════════
patches.append(('E3 Case21->22', 1,
crlf("""\
    Case 22 ' %ICMS
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota ICMS inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMS = Val(sVal)
        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(dblPICMS, 2)
        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMS, 2)"""),
crlf("""\
    Case 22 ' %ICMS
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota ICMS inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMS = Val(sVal)
        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), ".", ""), ",", ".")))
        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(dblPICMS, 2)
        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(curVICMS, 2)""")
))

# ══════════════════════════════════════════════════════════════
# E4 – Case 23 → 24 (cols 19→20, 21→22, 23→24, 20→21, 22→23)
# ══════════════════════════════════════════════════════════════
patches.append(('E4 Case23->24', 1,
crlf("""\
    Case 23 ' %RED BC
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPRedBC = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))
        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))
        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), ".", ""), ",", "."))
        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(dblPRedBC, 2)
        GridNotasItens.TextMatrix(iRow, 20) = FormatNumber(curVBC, 2)
        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMS, 2)"""),
crlf("""\
    Case 24 ' %RED BC
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPRedBC = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))
        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))
        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), ".", ""), ",", "."))
        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 24) = FormatNumber(dblPRedBC, 2)
        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(curVBC, 2)
        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(curVICMS, 2)""")
))

# ══════════════════════════════════════════════════════════════
# E5 – Case 25 → 26 (cols 24→25, 22→23, 25→26, 26→27)
# ══════════════════════════════════════════════════════════════
patches.append(('E5 Case25->26', 1,
crlf("""\
    Case 25 ' %ICMSST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMSST = Val(sVal)
        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 24), ".", ""), ",", ".")))
        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), ".", ""), ",", ".")))
        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(dblPICMSST, 2)
        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVICMSST, 2)"""),
crlf("""\
    Case 26 ' %ICMSST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPICMSST = Val(sVal)
        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), ".", ""), ",", ".")))
        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 23), ".", ""), ",", ".")))
        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(dblPICMSST, 2)
        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(curVICMSST, 2)""")
))

# ══════════════════════════════════════════════════════════════
# E6 – Case 27 → 28 (cols 19→20, 30→31, 25→26, 22→23, 27→28, 24→25, 26→27)
# ══════════════════════════════════════════════════════════════
patches.append(('E6 Case27->28', 1,
crlf("""\
    Case 27 ' MVA ST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then
            MsgBox "MVA inv\xe1lido (deve ser >= 0)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblMVA = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))
        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 30), ".", ""), ",", ".")))
        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))
        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), ".", ""), ",", "."))
        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), ".", ""), ",", ".")))
        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(dblMVA, 2)
        GridNotasItens.TextMatrix(iRow, 24) = FormatNumber(curVBCST, 2)
        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVICMSST, 2)"""),
crlf("""\
    Case 28 ' MVA ST
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then
            MsgBox "MVA inv\xe1lido (deve ser >= 0)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblMVA = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))
        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 31), ".", ""), ",", ".")))
        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))
        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 26), ".", ""), ",", "."))
        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 23), ".", ""), ",", ".")))
        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS
        If curVICMSST < 0 Then curVICMSST = 0
        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 28) = FormatNumber(dblMVA, 2)
        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(curVBCST, 2)
        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(curVICMSST, 2)""")
))

# ══════════════════════════════════════════════════════════════
# E7 – Case 28 → 29 (CST IPI)
# ══════════════════════════════════════════════════════════════
patches.append(('E7 Case28->29', 1,
crlf("""\
    Case 28 ' CST IPI
        If sVal = "" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then
            MsgBox "CST IPI deve ter 2 d\xedgitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET IPICST = '" & sVal & "', IPIcEnq = '999' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal"""),
crlf("""\
    Case 29 ' CST IPI
        If sVal = "" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then
            MsgBox "CST IPI deve ter 2 d\xedgitos!", vbInformation, "Aviso"
            Exit Sub
        End If
        dbData.Execute "UPDATE NotaFiscalItens SET IPICST = '" & sVal & "', IPIcEnq = '999' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, iCol) = sVal""")
))

# ══════════════════════════════════════════════════════════════
# E8 – Case 29 → 30 (%IPI) (cols 19→20, 29→30, 30→31)
# ══════════════════════════════════════════════════════════════
patches.append(('E8 Case29->30', 1,
crlf("""\
    Case 29 ' %IPI
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota IPI inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPIPI = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))
        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 29) = FormatNumber(dblPIPI, 2)
        GridNotasItens.TextMatrix(iRow, 30) = FormatNumber(curVIPI, 2)"""),
crlf("""\
    Case 30 ' %IPI
        sVal = Replace(Replace(sVal, ".", ""), ",", ".")
        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then
            MsgBox "Al\xedquota IPI inv\xe1lida (0 a 100)!", vbInformation, "Aviso"
            Exit Sub
        End If
        dblPIPI = Val(sVal)
        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))
        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))
        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)
        GridNotasItens.TextMatrix(iRow, 30) = FormatNumber(dblPIPI, 2)
        GridNotasItens.TextMatrix(iRow, 31) = FormatNumber(curVIPI, 2)""")
))

# ══════════════════════════════════════════════════════════════
# Aplicar patches
# ══════════════════════════════════════════════════════════════
for name, expected, old, new in patches:
    c = text.count(old)
    if c == expected:
        text = text.replace(old, new)
        print(f'  OK  {name}')
    else:
        print(f'ERRO  {name}: esperado {expected}, encontrado {c}')
        ok = False

if not ok:
    print('\nABORTADO — nenhuma alteracao salva')
    sys.exit(1)

out = text.encode('windows-1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(PATH, 'wb').write(out)
print('\nNFe_Completa.frm: salvo')
