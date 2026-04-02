data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# =============================================================================
# 1. GridNotasItens_Click — substituir loop For i=6 To 9 por Select Case
# =============================================================================
old_click = (
    b'Private Sub GridNotasItens_Click()\r\n'
    b'Dim i As Integer\r\n'
    b'\r\n'
    b'For i = 6 To 9\r\n'
    b'   If GridNotasItens.ColSel = i Then\r\n'
    b'      txtEdit.Move GridNotasItens.Left + GridNotasItens.CellLeft, GridNotasItens.Top + GridNotasItens.CellTop, GridNotasItens.CellWidth, GridNotasItens.CellHeight\r\n'
    b'      txtEdit.Text = GridNotasItens.TextMatrix(GridNotasItens.Row, GridNotasItens.Col)\r\n'
    b'      txtEdit.Visible = True\r\n'
    b'      txtEdit.SetFocus\r\n'
    b'      txtEdit.SelStart = 0\r\n'
    b'      txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    b'      iRow = GridNotasItens.Row\r\n'
    b'      iCol = GridNotasItens.Col\r\n'
    b'   End If\r\n'
    b'Next\r\n'
    b'End Sub\r\n'
)
new_click = (
    b'Private Sub GridNotasItens_Click()\r\n'
    b'Dim bEditavel As Boolean\r\n'
    b'bEditavel = False\r\n'
    b'\r\n'
    b'Select Case GridNotasItens.Col\r\n'
    b'    Case 2, 5, 6, 7, 8, 17, 19, 21, 23, 24, 25\r\n'
    b'        bEditavel = True\r\n'
    b'End Select\r\n'
    b'\r\n'
    b'If bEditavel And GridNotasItens.Row > 0 And GridNotasItens.TextMatrix(GridNotasItens.Row, 1) <> "" Then\r\n'
    b'    txtEdit.Move GridNotasItens.Left + GridNotasItens.CellLeft, GridNotasItens.Top + GridNotasItens.CellTop, GridNotasItens.CellWidth, GridNotasItens.CellHeight\r\n'
    b'    txtEdit.Text = GridNotasItens.TextMatrix(GridNotasItens.Row, GridNotasItens.Col)\r\n'
    b'    txtEdit.Visible = True\r\n'
    b'    txtEdit.SetFocus\r\n'
    b'    txtEdit.SelStart = 0\r\n'
    b'    txtEdit.SelLength = Len(txtEdit.Text)\r\n'
    b'    iRow = GridNotasItens.Row\r\n'
    b'    iCol = GridNotasItens.Col\r\n'
    b'End If\r\n'
    b'End Sub\r\n'
)
c = data.count(old_click)
print(f'1. GridNotasItens_Click: {c}')
if c == 1: data = data.replace(old_click, new_click)

# =============================================================================
# 2. txtEdit_LostFocus — reescrita completa
# =============================================================================
old_lost_start = b'Private Sub txtEdit_LostFocus()\r\n'
old_lost_end   = b'End Sub\r\nPrivate Sub Form_KeyDown'

idx_start = data.find(old_lost_start)
idx_end   = data.find(old_lost_end)
print(f'2. txtEdit_LostFocus start={idx_start} end={idx_end}')

if idx_start >= 0 and idx_end > idx_start:
    old_lost = data[idx_start : idx_end + len(b'End Sub\r\n')]

    new_lost = (
        b'Private Sub txtEdit_LostFocus()\r\n'
        b'Dim sVal      As String\r\n'
        b'Dim sItem     As String\r\n'
        b'Dim sCodProd  As String\r\n'
        b'Dim curVBC    As Currency\r\n'
        b'Dim curVICMS  As Currency\r\n'
        b'Dim curVBCST  As Currency\r\n'
        b'Dim curVICMSST As Currency\r\n'
        b'Dim curVIPI   As Currency\r\n'
        b'Dim curSubTot As Currency\r\n'
        b'Dim dblPICMS  As Double\r\n'
        b'Dim dblPICMSST As Double\r\n'
        b'Dim dblPRedBC As Double\r\n'
        b'Dim dblPIPI   As Double\r\n'
        b'Dim dblMVA    As Double\r\n'
        b'\r\n'
        b'txtEdit.Visible = False\r\n'
        b'sVal     = Trim(txtEdit.Text)\r\n'
        b'sItem    = GridNotasItens.TextMatrix(iRow, 1)\r\n'
        b'sCodProd = GridNotasItens.TextMatrix(iRow, 3)\r\n'
        b'\r\n'
        b'If sItem = "" Then Exit Sub\r\n'
        b'\r\n'
        b'Select Case iCol\r\n'
        b'\r\n'
        # --- Col 2: EAN ---
        b'    Case 2 \' EAN\r\n'
        b'        sVal = Replace(sVal, " ", "")\r\n'
        b'        If sVal <> "" Then\r\n'
        b'            If Not IsNumeric(sVal) Then\r\n'
        b'                MsgBox "EAN deve conter apenas d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'                Exit Sub\r\n'
        b'            End If\r\n'
        b'            If Len(sVal) <> 8 And Len(sVal) <> 13 Then\r\n'
        b'                MsgBox "EAN deve ter 8 ou 13 d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'                Exit Sub\r\n'
        b'            End If\r\n'
        b'        End If\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET EAN = \'" & sVal & "\' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        dbData.Execute "UPDATE Produtos SET EAN = \'" & sVal & "\', COD_BARRA = \'" & sVal & "\' WHERE CODIGO = " & Val(sCodProd)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 5: UND ---
        b'    Case 5 \' UND\r\n'
        b'        If sVal = "" Then\r\n'
        b'            MsgBox "Unidade n\xe3o pode ser vazia!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        If Len(sVal) > 2 Then\r\n'
        b'            MsgBox "Unidade deve ter no m\xe1ximo 2 caracteres!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        sVal = UCase(sVal)\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET UnidadeComercial = \'" & sVal & "\' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        dbData.Execute "UPDATE Produtos SET unid_medida = \'" & sVal & "\' WHERE CODIGO = " & Val(sCodProd)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 6: NCM ---
        b'    Case 6 \' NCM\r\n'
        b'        sVal = Replace(sVal, ".", "")\r\n'
        b'        If sVal <> "" Then\r\n'
        b'            If Len(sVal) <> 8 Or Not IsNumeric(sVal) Then\r\n'
        b'                MsgBox "NCM deve ter 8 d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'                Exit Sub\r\n'
        b'            End If\r\n'
        b'        End If\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET NCM = \'" & sVal & "\' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        dbData.Execute "UPDATE Produtos SET NCM = \'" & sVal & "\' WHERE CODIGO = " & Val(sCodProd)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 7: CFOP ---
        b'    Case 7 \' CFOP\r\n'
        b'        If sVal = "" Or Len(sVal) <> 4 Or Not IsNumeric(sVal) Then\r\n'
        b'            MsgBox "CFOP deve ter 4 d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET CFOP = " & Val(sVal) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 8: CST ---
        b'    Case 8 \' CST\r\n'
        b'        If sVal = "" Or Len(sVal) <> 3 Then\r\n'
        b'            MsgBox "CST deve ter 3 d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET CST = \'" & sVal & "\' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 17: %ICMS ---
        b'    Case 17 \' %ICMS\r\n'
        b'        sVal = Replace(sVal, ",", ".")\r\n'
        b'        If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Or CDbl(sVal) > 100 Then\r\n'
        b'            MsgBox "Al\xedquota ICMS inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dblPICMS = CDbl(sVal)\r\n'
        b'        curVBC   = CCur(Replace(GridNotasItens.TextMatrix(iRow, 16), ",", "."))\r\n'
        b'        curVICMS = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET pICMS = " & FSQL(dblPICMS, 4) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 17) = FormatNumber(dblPICMS, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)\r\n'
        b'\r\n'
        # --- Col 19: %REDBC ---
        b'    Case 19 \' %RED BC\r\n'
        b'        sVal = Replace(sVal, ",", ".")\r\n'
        b'        If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Or CDbl(sVal) > 100 Then\r\n'
        b'            MsgBox "Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dblPRedBC = CDbl(sVal)\r\n'
        b'        curSubTot = CCur(Replace(GridNotasItens.TextMatrix(iRow, 15), ",", "."))\r\n'
        b'        curVBC    = CCur(Format(curSubTot * (1 - dblPRedBC / 100), "0.00"))\r\n'
        b'        dblPICMS  = CDbl(Replace(GridNotasItens.TextMatrix(iRow, 17), ",", "."))\r\n'
        b'        curVICMS  = CCur(Format(curVBC * dblPICMS / 100, "0.00"))\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET pRedBC = " & FSQL(dblPRedBC, 4) & ", vBC = " & FSQL(curVBC, 2) & ", vICMS = " & FSQL(curVICMS, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 19) = FormatNumber(dblPRedBC, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 16) = FormatNumber(curVBC, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 18) = FormatNumber(curVICMS, 2)\r\n'
        b'\r\n'
        # --- Col 21: %ICMSST ---
        b'    Case 21 \' %ICMSST\r\n'
        b'        sVal = Replace(sVal, ",", ".")\r\n'
        b'        If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Or CDbl(sVal) > 100 Then\r\n'
        b'            MsgBox "Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dblPICMSST  = CDbl(sVal)\r\n'
        b'        curVBCST    = CCur(Replace(GridNotasItens.TextMatrix(iRow, 20), ",", "."))\r\n'
        b'        curVICMS    = CCur(Replace(GridNotasItens.TextMatrix(iRow, 18), ",", "."))\r\n'
        b'        curVICMSST  = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\r\n'
        b'        If curVICMSST < 0 Then curVICMSST = 0\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET pICMSST = " & FSQL(dblPICMSST, 4) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(dblPICMSST, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)\r\n'
        b'\r\n'
        # --- Col 23: MVA ST ---
        b'    Case 23 \' MVA ST\r\n'
        b'        sVal = Replace(sVal, ",", ".")\r\n'
        b'        If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Then\r\n'
        b'            MsgBox "MVA inv\xe1lido (deve ser >= 0)!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dblMVA      = CDbl(sVal)\r\n'
        b'        curSubTot   = CCur(Replace(GridNotasItens.TextMatrix(iRow, 15), ",", "."))\r\n'
        b'        curVIPI     = CCur(Replace(GridNotasItens.TextMatrix(iRow, 26), ",", "."))\r\n'
        b'        curVBCST    = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), "0.00"))\r\n'
        b'        dblPICMSST  = CDbl(Replace(GridNotasItens.TextMatrix(iRow, 21), ",", "."))\r\n'
        b'        curVICMS    = CCur(Replace(GridNotasItens.TextMatrix(iRow, 18), ",", "."))\r\n'
        b'        curVICMSST  = CCur(Format(curVBCST * dblPICMSST / 100, "0.00")) - curVICMS\r\n'
        b'        If curVICMSST < 0 Then curVICMSST = 0\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET pMVAST = " & FSQL(dblMVA, 4) & ", vBCST = " & FSQL(curVBCST, 2) & ", vICMSST = " & FSQL(curVICMSST, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(dblMVA, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 20) = FormatNumber(curVBCST, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(curVICMSST, 2)\r\n'
        b'\r\n'
        # --- Col 24: CST IPI ---
        b'    Case 24 \' CST IPI\r\n'
        b'        If sVal = "" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then\r\n'
        b'            MsgBox "CST IPI deve ter 2 d\xedgitos!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET IPICST = \'" & sVal & "\', IPIcEnq = \'999\' WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n'
        b'\r\n'
        # --- Col 25: %IPI ---
        b'    Case 25 \' %IPI\r\n'
        b'        sVal = Replace(sVal, ",", ".")\r\n'
        b'        If Not IsNumeric(sVal) Or CDbl(sVal) < 0 Or CDbl(sVal) > 100 Then\r\n'
        b'            MsgBox "Al\xedquota IPI inv\xe1lida (0 a 100)!", vbInformation, "Aviso"\r\n'
        b'            Exit Sub\r\n'
        b'        End If\r\n'
        b'        dblPIPI   = CDbl(sVal)\r\n'
        b'        curSubTot = CCur(Replace(GridNotasItens.TextMatrix(iRow, 15), ",", "."))\r\n'
        b'        curVIPI   = CCur(Format(curSubTot * dblPIPI / 100, "0.00"))\r\n'
        b'        dbData.Execute "UPDATE NotaFiscalItens SET IPIpIPI = " & FSQL(dblPIPI, 4) & ", IPIvIPI = " & FSQL(curVIPI, 2) & " WHERE CodigoNota = " & Val(txtCodNota.Text) & " AND ITEM = " & Val(sItem)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(dblPIPI, 2)\r\n'
        b'        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(curVIPI, 2)\r\n'
        b'\r\n'
        b'End Select\r\n'
        b'\r\n'
        b'AtualizarTotaisNota\r\n'
        b'End Sub\r\n'
    )
    data = data[:idx_start] + new_lost + data[idx_start + len(old_lost):]
    print('   txtEdit_LostFocus substituido.')
else:
    print('   ERRO: nao encontrou os limites de txtEdit_LostFocus!')

# =============================================================================
# 3. AtualizarGrid_Itens — simplificar (individual UPDATEs agora no LostFocus)
# =============================================================================
old_agrid = (
    b'Private Sub AtualizarGrid_Itens()\r\n'
    b'Dim i As Integer\r\n'
    b'   \r\n'
    b'For i = 1 To GridNotasItens.rows - 1\r\n'
    b'   If GridNotasItens.TextMatrix(i, 1) <> "" Then  \'vValorIcmsLinha\r\n'
    b'      dbData.Execute "UPDATE NotaFiscalItens SET CFOP = " & GridNotasItens.TextMatrix(i, 7) & ", CST = \'" & GridNotasItens.TextMatrix(i, 8) & "\', NCM = \'" & GridNotasItens.TextMatrix(i, 6) & "\', pICMS = " & FSQL(GridNotasItens.TextMatrix(i, 9), 2) & ", vICMS = " & FSQL(GridNotasItens.TextMatrix(i, 10), 2) & "  WHERE CodigoNota = " & txtCodNota.Text & " AND ITEM = " & GridNotasItens.TextMatrix(i, 1) & ""\r\n'
    b'      dbData.Execute "UPDATE TbNFCe_Itens SET CodNcm = \'" & GridNotasItens.TextMatrix(i, 6) & "\' WHERE IDProduto = " & GridNotasItens.TextMatrix(i, 3) & ""\r\n'
    b'      dbData.Execute "UPDATE Produtos SET NCM = \'" & GridNotasItens.TextMatrix(i, 6) & "\' WHERE CODIGO = " & GridNotasItens.TextMatrix(i, 3) & ""\r\n'
    b'   End If\r\n'
    b'Next\r\n'
    b'\'Call MostrarValorBaseICMS\r\n'
    b'Call AtualizarValorICMS\r\n'
    b'Call CalcularIPI\r\n'
    b'End Sub\r\n'
)
new_agrid = (
    b'Private Sub AtualizarGrid_Itens()\r\n'
    b'AtualizarTotaisNota\r\n'
    b'End Sub\r\n'
)
c = data.count(old_agrid)
print(f'3. AtualizarGrid_Itens: {c}')
if c == 1: data = data.replace(old_agrid, new_agrid)

# Normalizar CRLF
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
