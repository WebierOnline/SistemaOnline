import sys

with open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'rb') as f:
    data = f.read()

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n')
text = data.decode('cp1252')

errors = []

def replace_one(text, old, new, label):
    cnt = text.count(old)
    if cnt != 1:
        errors.append(f'{label}: encontrado {cnt} ocorrencias (esperado 1)')
        return text
    return text.replace(old, new)

# ============================================================
# 1. Exibir_Itens — adiciona campos IBS/CBS/IS ao SELECT
# ============================================================
old1 = (
    'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, " & _\n'
    '       "ValorUnitarioComercializacao, QuantidadeComercial, ValorTotalBruto, " & _\n'
    '       "ValorFrete, ValorSeguro, ValorOutros, ValorDesconto, " & _\n'
    '       "vBC, pICMS, vICMS, pRedBC, " & _\n'
    '       "vBCST, pICMSST, vICMSST, pMVAST, " & _\n'
    '       "IPICST, IPIpIPI, IPIvIPI " & _\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)'
)
new1 = (
    'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, " & _\n'
    '       "IBSCBS_CST, IBS_vIBS, CBS_vCBS, IS_vIS, " & _\n'
    '       "ValorUnitarioComercializacao, QuantidadeComercial, ValorTotalBruto, " & _\n'
    '       "ValorFrete, ValorSeguro, ValorOutros, ValorDesconto, " & _\n'
    '       "vBC, pICMS, vICMS, pRedBC, " & _\n'
    '       "vBCST, pICMSST, vICMSST, pMVAST, " & _\n'
    '       "IPICST, IPIpIPI, IPIvIPI " & _\n'
    '       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)'
)
text = replace_one(text, old1, new1, 'Exibir_Itens SELECT')

# ============================================================
# 2. FormatarGridItensNota — substituicao completa do sub
# ============================================================
old2 = \
"""Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim j As Integer

   With GridNotasItens
      .Visible = False
      .Redraw = False

      .Clear
      .Cols = 27
      .rows = 2

      'Colunas fixas (sempre visiveis)
      .ColWidth(0) = 300    'indicador de linha
      .ColWidth(1) = 0      'No.
      .ColWidth(2) = 1500   'EAN
      .ColWidth(3) = 0      'COD. (oculto)
      .ColWidth(4) = 3500   'DESCRICAO
      .ColWidth(5) = 450    'UND
      .ColWidth(6) = 900    'NCM
      .ColWidth(7) = 600    'CFOP
      .ColWidth(8) = 500    'CST
      .ColWidth(9) = 850    'VALOR
      .ColWidth(10) = 850   'QTDE
      .ColWidth(11) = 800   'FRETE
      .ColWidth(12) = 900   'SEGURO
      .ColWidth(13) = 900   'OUTROS
      .ColWidth(14) = 800   'DESC.
      .ColWidth(15) = 1050  'TOTAL
      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)
      .ColWidth(16) = 0     'BC ICMS
      .ColWidth(17) = 0     '%ICMS
      .ColWidth(18) = 0     'ICMS
      .ColWidth(19) = 0     '%RED BC
      .ColWidth(20) = 0     'BC ST
      .ColWidth(21) = 0     '%ICMSST
      .ColWidth(22) = 0     'ICMSST
      .ColWidth(23) = 0     'MVA ST
      .ColWidth(24) = 0     '%IPI
      .ColWidth(25) = 0     'IPI
      .ColWidth(26) = 0     'cEnq

      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "EAN"
      .TextMatrix(0, 3) = "C\xd3D."
      .TextMatrix(0, 4) = "DESCRI\xc7\xc3O"
      .TextMatrix(0, 5) = "UND"
      .TextMatrix(0, 6) = "NCM"
      .TextMatrix(0, 7) = "CFOP"
      .TextMatrix(0, 8) = "CST"
      .TextMatrix(0, 9) = "VALOR"
      .TextMatrix(0, 10) = "QTDE"
      .TextMatrix(0, 11) = "FRETE"
      .TextMatrix(0, 12) = "SEGURO"
      .TextMatrix(0, 13) = "OUTROS"
      .TextMatrix(0, 14) = "DESC."
      .TextMatrix(0, 15) = "TOTAL"
      .TextMatrix(0, 16) = "BC ICMS"
      .TextMatrix(0, 17) = "%ICMS"
      .TextMatrix(0, 18) = "ICMS"
      .TextMatrix(0, 19) = "%RED BC"
      .TextMatrix(0, 20) = "BC ST"
      .TextMatrix(0, 21) = "%ICMSST"
      .TextMatrix(0, 22) = "ICMSST"
      .TextMatrix(0, 23) = "MVA ST"
      .TextMatrix(0, 24) = "CST IPI"
      .TextMatrix(0, 25) = "%IPI"
      .TextMatrix(0, 26) = "IPI"

      'Cabecalho em negrito e centralizado
      For i = 0 To .Cols - 1
         .Col = i: .Row = 0
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next i

      'Alinhamento: texto esquerda (0-8), numeros direita (9-26)
      For i = 0 To 8
         .ColAlignment(i) = 1
      Next i
      For i = 9 To 26
         .ColAlignment(i) = 6
      Next i

      i = 1
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("ITEM"), "000")
            .TextMatrix(.rows - 1, 2) = rTabela("EAN")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("CodigoProduto"), "00000")
            .TextMatrix(.rows - 1, 4) = rTabela("NomeProduto")
            .TextMatrix(.rows - 1, 5) = rTabela("UnidadeComercial")
            .TextMatrix(.rows - 1, 6) = rTabela("NCM")
            .TextMatrix(.rows - 1, 7) = rTabela("CFOP")
            .TextMatrix(.rows - 1, 8) = rTabela("CST")
            .TextMatrix(.rows - 1, 9) = FormatNumber(rTabela("ValorUnitarioComercializacao"), 2)
            If rTabela("UnidadeComercial") = "KG" Or rTabela("UnidadeComercial") = "GR" Or rTabela("UnidadeComercial") = "MG" Then
                .TextMatrix(.rows - 1, 10) = Format(rTabela("QuantidadeComercial"), ocPESO)
            Else
                .TextMatrix(.rows - 1, 10) = Format(rTabela("QuantidadeComercial"), "###,###,##0")
            End If
            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("ValorFrete"), 2)
            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("ValorSeguro"), 2)
            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela("ValorOutros"), 2)
            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("ValorDesconto"), 2)
            .TextMatrix(.rows - 1, 15) = FormatNumber(rTabela("ValorTotalBruto"), 2)
            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela("vBC"), 2)
            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela("pICMS"), 2)
            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela("vICMS"), 2)
            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela("pRedBC"), 2)
            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela("vBCST"), 2)
            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela("pICMSST"), 2)
            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela("vICMSST"), 2)
            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela("pMVAST"), 2)
            .TextMatrix(.rows - 1, 24) = rTabela("IPICST")
            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("IPIpIPI"), 2)
            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela("IPIvIPI"), 2)

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If

      .rows = .rows - 1

      'Numero da linha no col 0
      For i = 1 To .rows - 1
         .TextMatrix(i, 0) = i
      Next i

      'EAN em negrito
      For i = 1 To .rows - 1
         .Row = i: .Col = 2: .CellFontBold = True
      Next i

      'COD. em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'TOTAL em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 15
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'Colunas edit\xe1veis em amarelo claro
      Dim colEdit As Variant
      For Each colEdit In Array(2, 5, 6, 7, 8, 17, 19, 21, 23, 24, 25)
         For i = 1 To .rows - 1
            .Row = i: .Col = colEdit
            .CellBackColor = &HC8FFFF
         Next i
      Next colEdit

      GridNotasItens.Col = 0
      .Visible = True
      .Redraw = True
   End With
End Sub"""

new2 = \
"""Sub FormatarGridItensNota(rTabela As ADODB.Recordset)
   Dim i As Integer
   Dim j As Integer

   With GridNotasItens
      .Visible = False
      .Redraw = False

      .Clear
      .Cols = 31
      .rows = 2

      'Colunas fixas (sempre visiveis)
      .ColWidth(0) = 300    'indicador de linha
      .ColWidth(1) = 0      'No.
      .ColWidth(2) = 1500   'EAN
      .ColWidth(3) = 0      'COD. (oculto)
      .ColWidth(4) = 3500   'DESCRICAO
      .ColWidth(5) = 450    'UND
      .ColWidth(6) = 900    'NCM
      .ColWidth(7) = 600    'CFOP
      .ColWidth(8) = 500    'CST
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
      .ColWidth(30) = 0     'cEnq

      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "EAN"
      .TextMatrix(0, 3) = "C\xd3D."
      .TextMatrix(0, 4) = "DESCRI\xc7\xc3O"
      .TextMatrix(0, 5) = "UND"
      .TextMatrix(0, 6) = "NCM"
      .TextMatrix(0, 7) = "CFOP"
      .TextMatrix(0, 8) = "CST"
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
      .TextMatrix(0, 30) = "IPI"

      'Cabecalho em negrito e centralizado
      For i = 0 To .Cols - 1
         .Col = i: .Row = 0
         .CellFontBold = True
         .CellAlignment = flexAlignCenterCenter
      Next i

      'Alinhamento: texto esquerda (0-8), numeros direita (9-30)
      For i = 0 To 8
         .ColAlignment(i) = 1
      Next i
      For i = 9 To 30
         .ColAlignment(i) = 6
      Next i

      i = 1
      If Not rTabela Is Nothing Then
         Do While Not rTabela.EOF
            .TextMatrix(.rows - 1, 1) = Format(rTabela("ITEM"), "000")
            .TextMatrix(.rows - 1, 2) = rTabela("EAN")
            .TextMatrix(.rows - 1, 3) = Format(rTabela("CodigoProduto"), "00000")
            .TextMatrix(.rows - 1, 4) = rTabela("NomeProduto")
            .TextMatrix(.rows - 1, 5) = rTabela("UnidadeComercial")
            .TextMatrix(.rows - 1, 6) = rTabela("NCM")
            .TextMatrix(.rows - 1, 7) = rTabela("CFOP")
            .TextMatrix(.rows - 1, 8) = rTabela("CST")
            .TextMatrix(.rows - 1, 9) = rTabela("IBSCBS_CST")
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
            .TextMatrix(.rows - 1, 30) = FormatNumber(rTabela("IPIvIPI"), 2)

            rTabela.MoveNext
            .rows = .rows + 1
            i = i + 1
         Loop
      End If

      .rows = .rows - 1

      'Numero da linha no col 0
      For i = 1 To .rows - 1
         .TextMatrix(i, 0) = i
      Next i

      'EAN em negrito
      For i = 1 To .rows - 1
         .Row = i: .Col = 2: .CellFontBold = True
      Next i

      'COD. em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 3
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'TOTAL em destaque
      For i = 1 To .rows - 1
         .Row = i: .Col = 19
         .CellForeColor = &HC0&
         .CellFontBold = True
      Next i

      'Colunas edit\xe1veis em amarelo claro
      Dim colEdit As Variant
      For Each colEdit In Array(2, 5, 6, 7, 8, 21, 23, 25, 27, 28, 29)
         For i = 1 To .rows - 1
            .Row = i: .Col = colEdit
            .CellBackColor = &HC8FFFF
         Next i
      Next colEdit

      'Colunas reforma tributaria em azul claro
      Dim colRef As Variant
      For Each colRef In Array(9, 10, 11, 12)
         For i = 1 To .rows - 1
            .Row = i: .Col = colRef
            .CellBackColor = &HFFFFF0
         Next i
      Next colRef

      GridNotasItens.Col = 0
      .Visible = True
      .Redraw = True
   End With
End Sub"""

text = replace_one(text, old2, new2, 'FormatarGridItensNota')

# ============================================================
# 3. AplicarVisibilidadeGridItens — novos checkboxes + renumber
# ============================================================
old3 = \
"""Sub AplicarVisibilidadeGridItens()
   If GridNotasItens.Cols < 27 Then Exit Sub
   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)
   Dim bICMS As Boolean
   bICMS = (Left(cboFinalidade.Text, 1) = "4")
   GridNotasItens.ColWidth(16) = IIf(bICMS, 850, 0)  'BC ICMS
   GridNotasItens.ColWidth(17) = IIf(bICMS, 850, 0)  '%ICMS
   GridNotasItens.ColWidth(18) = IIf(bICMS, 850, 0)  'ICMS

   '%RedBC: chkpRedBC
   GridNotasItens.ColWidth(19) = IIf(chkpRedBC.Value = 1, 700, 0)

   'Grupo ICMSST: chkICMSST
   Dim bST As Boolean
   bST = (chkICMSST.Value = 1)
   GridNotasItens.ColWidth(20) = IIf(bST, 850, 0)  'BC ST
   GridNotasItens.ColWidth(21) = IIf(bST, 900, 0)  '%ICMSST
   GridNotasItens.ColWidth(22) = IIf(bST, 850, 0)  'ICMSST
   GridNotasItens.ColWidth(23) = IIf(bST, 850, 0)  'MVA ST

   'Grupo IPI: chkIPI
   Dim bIPI As Boolean
   bIPI = (chkIPI.Value = 1)
   GridNotasItens.ColWidth(24) = IIf(bIPI, 850, 0)  '%IPI
   GridNotasItens.ColWidth(25) = IIf(bIPI, 850, 0)  'IPI
   GridNotasItens.ColWidth(26) = IIf(bIPI, 850, 0)  'cEnq
End Sub"""

new3 = \
"""Sub AplicarVisibilidadeGridItens()
   If GridNotasItens.Cols < 31 Then Exit Sub
   'Reforma tributaria: chkReforma
   Dim bReforma As Boolean
   bReforma = (chkReforma.Value = 1)
   GridNotasItens.ColWidth(9)  = IIf(bReforma, 700, 0)   'CST IBS/CBS
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
End Sub"""

text = replace_one(text, old3, new3, 'AplicarVisibilidadeGridItens')

# ============================================================
# 4. GridNotasItens_Click — atualiza colunas editaveis
# ============================================================
old4 = '    Case 2, 5, 6, 7, 8, 17, 19, 21, 23, 24, 25\n        bEditavel = True'
new4 = '    Case 2, 5, 6, 7, 8, 21, 23, 25, 27, 28, 29\n        bEditavel = True'
text = replace_one(text, old4, new4, 'GridNotasItens_Click editaveis')

# ============================================================
# 5. txtEdit_LostFocus — renumerar todos os casos e colunas
# ============================================================

# Case 5 (UND): col 10 -> 14
old5a = (
    '        sQtdAtual = GridNotasItens.TextMatrix(iRow, 10)\n'
    '        If sVal = "KG" Or sVal = "GR" Or sVal = "MG" Then\n'
    '            GridNotasItens.TextMatrix(iRow, 10) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), ocPESO)\n'
    '        Else\n'
    '            GridNotasItens.TextMatrix(iRow, 10) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), "###,###,##0")\n'
    '        End If'
)
new5a = (
    '        sQtdAtual = GridNotasItens.TextMatrix(iRow, 14)\n'
    '        If sVal = "KG" Or sVal = "GR" Or sVal = "MG" Then\n'
    '            GridNotasItens.TextMatrix(iRow, 14) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), ocPESO)\n'
    '        Else\n'
    '            GridNotasItens.TextMatrix(iRow, 14) = Format(Val(Replace(Replace(sQtdAtual, ".", ""), ",", ".")), "###,###,##0")\n'
    '        End If'
)
text = replace_one(text, old5a, new5a, 'txtEdit Case5 UND col10->14')

# Case 17 -> 21 (%ICMS)
old5b = (
    '    Case 17 \' %ICMS\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota ICMS inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPICMS = Val(sVal)\n'
    '        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 16), ".", ""), ",", ".")))\n'
    '        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 17) = FormatNumber(dblPICMS, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)'
)
new5b = (
    '    Case 21 \' %ICMS\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota ICMS inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPICMS = Val(sVal)\n'
    '        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))\n'
    '        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(dblPICMS, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMS, 2)'
)
text = replace_one(text, old5b, new5b, 'txtEdit Case17->21 %ICMS')

# Case 19 -> 23 (%RED BC)
old5c = (
    '    Case 19 \' %RED BC\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPRedBC = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))\n'
    '        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))\n'
    '        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 17), ".", ""), ",", "."))\n'
    '        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 19) = FormatNumber(dblPRedBC, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 16) = FormatNumber(curVBC, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)'
)
new5c = (
    '    Case 23 \' %RED BC\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPRedBC = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))\n'
    '        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))\n'
    '        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), ".", ""), ",", "."))\n'
    '        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(dblPRedBC, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 20) = FormatNumber(curVBC, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMS, 2)'
)
text = replace_one(text, old5c, new5c, 'txtEdit Case19->23 %REDBC')

# Case 21 -> 25 (%ICMSST)
old5d = (
    '    Case 21 \' %ICMSST\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPICMSST = Val(sVal)\n'
    '        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), ".", ""), ",", ".")))\n'
    '        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 18), ".", ""), ",", ".")))\n'
    '        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\n'
    '        If curVICMSST < 0 Then curVICMSST = 0\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(dblPICMSST, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)'
)
new5d = (
    '    Case 25 \' %ICMSST\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPICMSST = Val(sVal)\n'
    '        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 24), ".", ""), ",", ".")))\n'
    '        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), ".", ""), ",", ".")))\n'
    '        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\n'
    '        If curVICMSST < 0 Then curVICMSST = 0\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(dblPICMSST, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVICMSST, 2)'
)
text = replace_one(text, old5d, new5d, 'txtEdit Case21->25 %ICMSST')

# Case 23 -> 27 (MVA ST)
old5e = (
    '    Case 23 \' MVA ST\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then\n'
    '            MsgBox "MVA inv\xe1lido (deve ser >= 0)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblMVA = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))\n'
    '        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 26), ".", ""), ",", ".")))\n'
    '        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))\n'
    '        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), ".", ""), ",", "."))\n'
    '        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 18), ".", ""), ",", ".")))\n'
    '        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\n'
    '        If curVICMSST < 0 Then curVICMSST = 0\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(dblMVA, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 20) = FormatNumber(curVBCST, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)'
)
new5e = (
    '    Case 27 \' MVA ST\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then\n'
    '            MsgBox "MVA inv\xe1lido (deve ser >= 0)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblMVA = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))\n'
    '        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 30), ".", ""), ",", ".")))\n'
    '        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))\n'
    '        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), ".", ""), ",", "."))\n'
    '        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), ".", ""), ",", ".")))\n'
    '        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\n'
    '        If curVICMSST < 0 Then curVICMSST = 0\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(dblMVA, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 24) = FormatNumber(curVBCST, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVICMSST, 2)'
)
text = replace_one(text, old5e, new5e, 'txtEdit Case23->27 MVAST')

# Case 24 -> 28 (CST IPI)
old5f = "    Case 24 ' CST IPI"
new5f = "    Case 28 ' CST IPI"
text = replace_one(text, old5f, new5f, 'txtEdit Case24->28 CSTIPI')

# Case 25 -> 29 (%IPI)
old5g = (
    '    Case 25 \' %IPI\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota IPI inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPIPI = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 15), ".", ""), ",", ".")))\n'
    '        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(dblPIPI, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVIPI, 2)'
)
new5g = (
    '    Case 29 \' %IPI\n'
    '        sVal = Replace(Replace(sVal, ".", ""), ",", ".")\n'
    '        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\n'
    '            MsgBox "Al\xedquota IPI inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\n'
    '            Exit Sub\n'
    '        End If\n'
    '        dblPIPI = Val(sVal)\n'
    '        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 19), ".", ""), ",", ".")))\n'
    '        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))\n'
    '        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\n'
    '        GridNotasItens.TextMatrix(iRow, 29) = FormatNumber(dblPIPI, 2)\n'
    '        GridNotasItens.TextMatrix(iRow, 30) = FormatNumber(curVIPI, 2)'
)
text = replace_one(text, old5g, new5g, 'txtEdit Case25->29 %IPI')

# ============================================================
# 6. Inserir handlers chkReforma_Click, chkSeguro_Click, chkOutros_Click
#    Insere apos chkICMSST_Click
# ============================================================
old6 = \
"""Sub chkICMSST_Click()
    If bSupressChkEvents Then Exit Sub
    AplicarVisibilidadeGridItens
    RecalcularItensNota
End Sub"""

new6 = \
"""Sub chkICMSST_Click()
    If bSupressChkEvents Then Exit Sub
    AplicarVisibilidadeGridItens
    RecalcularItensNota
End Sub

Sub chkReforma_Click()
    AplicarVisibilidadeGridItens
End Sub

Sub chkSeguro_Click()
    AplicarVisibilidadeGridItens
End Sub

Sub chkOutros_Click()
    AplicarVisibilidadeGridItens
End Sub"""

text = replace_one(text, old6, new6, 'chkICMSST_Click + novos handlers')

# ============================================================
# Verificar erros e gravar
# ============================================================
if errors:
    print('ERROS:')
    for e in errors:
        print(' -', e)
    sys.exit(1)

out = text.encode('cp1252')
out = out.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
with open(r'C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm', 'wb') as f:
    f.write(out)

print('OK - patch aplicado com sucesso')
print(f'Tamanho final: {len(out)} bytes')
