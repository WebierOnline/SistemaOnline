data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()

# =============================================================================
# 1. Reescrever cboNatureza_GotFocus para usar CarregarNaturezas
# =============================================================================
idx_start = data.find(b'Private Sub cboNatureza_GotFocus()')
idx_end   = data.find(b'\r\nEnd Sub\r\n', idx_start) + len(b'\r\nEnd Sub\r\n')
old_gotfocus = data[idx_start:idx_end]

new_gotfocus = (
    b'Private Sub cboNatureza_GotFocus()\r\n'
    b'Dim itemAtual As String\r\n'
    b'itemAtual = cboNatureza.Text\r\n'
    b'CarregarNaturezas\r\n'
    b'cboNatureza.Text = itemAtual\r\n'
    b'SelectControl cboNatureza\r\n'
    b'moCombo.AttachTo cboNatureza\r\n'
    b'End Sub\r\n'
)
data = data[:idx_start] + new_gotfocus + data[idx_end:]
print('1. cboNatureza_GotFocus reescrito')

# =============================================================================
# 2. Inserir CarregarNaturezas + AjustarNaturezaPadrao antes de cboDestOperacao_Change
# =============================================================================
old_change_anchor = b'Private Sub cboDestOperacao_Change()\r\n'
new_subs = (
    b'Private Sub CarregarNaturezas()\r\n'
    b'Dim r       As ADODB.Recordset\r\n'
    b'Dim nMin    As Long\r\n'
    b'Dim nMax    As Long\r\n'
    b'Dim sAvDev  As String\r\n'
    b'Dim sWhere  As String\r\n'
    b'\r\n'
    b'cboNatureza.Clear\r\n'
    b'\r\n'
    b'\' Prefixo: TipoNota + DestOperacao\r\n'
    b'If Left(cboTipoNota.Text, 1) = "1" Then  \' SAIDA\r\n'
    b'    Select Case Left(cboDestOperacao.Text, 1)\r\n'
    b'        Case "2": nMin = 6000: nMax = 6999\r\n'
    b'        Case "3": nMin = 7000: nMax = 7999\r\n'
    b'        Case Else: nMin = 5000: nMax = 5999\r\n'
    b'    End Select\r\n'
    b'Else  \' ENTRADA\r\n'
    b'    Select Case Left(cboDestOperacao.Text, 1)\r\n'
    b'        Case "2": nMin = 2000: nMax = 2999\r\n'
    b'        Case "3": nMin = 3000: nMax = 3999\r\n'
    b'        Case Else: nMin = 1000: nMax = 1999\r\n'
    b'    End Select\r\n'
    b'End If\r\n'
    b'\r\n'
    b'\' Avulsos de devolucao fora do range xx2xx\r\n'
    b'Select Case nMin\r\n'
    b'    Case 1000: sAvDev = "1410,1411,1902,1904,1906"\r\n'
    b'    Case 2000: sAvDev = "2902"\r\n'
    b'    Case 5000: sAvDev = "5411,5413,5553,5556,5902,5906,5913"\r\n'
    b'    Case 6000: sAvDev = "6411,6413,6553,6556,6902,6906"\r\n'
    b'    Case Else: sAvDev = ""\r\n'
    b'End Select\r\n'
    b'\r\n'
    b'\' Sub-filtro de finalidade\r\n'
    b'Select Case Left(cboFinalidade.Text, 1)\r\n'
    b'    Case "3"\r\n'
    b'        sWhere = " WHERE (CodigoNatureza BETWEEN " & (nMin + 600) & " AND " & (nMin + 699) & ")" & _\r\n'
    b'                 "    OR CodigoNatureza = " & (nMin + 933)\r\n'
    b'    Case "4"\r\n'
    b'        sWhere = " WHERE (CodigoNatureza BETWEEN " & (nMin + 200) & " AND " & (nMin + 299) & ")"\r\n'
    b'        If sAvDev <> "" Then sWhere = sWhere & " OR CodigoNatureza IN (" & sAvDev & ")"\r\n'
    b'    Case Else\r\n'
    b'        sWhere = " WHERE CodigoNatureza BETWEEN " & nMin & " AND " & nMax & _\r\n'
    b'                 " AND NOT (CodigoNatureza BETWEEN " & (nMin + 200) & " AND " & (nMin + 299) & ")" & _\r\n'
    b'                 " AND NOT (CodigoNatureza BETWEEN " & (nMin + 600) & " AND " & (nMin + 699) & ")" & _\r\n'
    b'                 " AND CodigoNatureza <> " & (nMin + 933)\r\n'
    b'        If sAvDev <> "" Then sWhere = sWhere & " AND CodigoNatureza NOT IN (" & sAvDev & ")"\r\n'
    b'End Select\r\n'
    b'\r\n'
    b'sSQL = "SELECT CodigoNatureza FROM NaturezaOperacaoNF" & sWhere & " ORDER BY CodigoNatureza"\r\n'
    b'Set r = dbData.OpenRecordset(sSQL)\r\n'
    b'Do While Not r.EOF\r\n'
    b'    cboNatureza.AddItem r("CodigoNatureza")\r\n'
    b'    r.MoveNext\r\n'
    b'Loop\r\n'
    b'If r.State <> 0 Then r.Close\r\n'
    b'Set r = Nothing\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub AjustarNaturezaPadrao()\r\n'
    b'\' Auto-ajuste quando ENTRADA: Finalidade=3 e DestOperacao=1\r\n'
    b'If Left(cboTipoNota.Text, 1) = "0" Then\r\n'
    b'    cboFinalidade.Text    = "3 - NFe DE AJUSTE"\r\n'
    b'    cboDestOperacao.Text  = "1 - Opera\xe7\xe3o Interna"\r\n'
    b'End If\r\n'
    b'\r\n'
    b'CarregarNaturezas\r\n'
    b'\r\n'
    b'\' Default especifico por combinacao\r\n'
    b'Dim sTipo As String, sFin As String, sDest As String\r\n'
    b'sTipo = Left(cboTipoNota.Text, 1)\r\n'
    b'sFin  = Left(cboFinalidade.Text, 1)\r\n'
    b'sDest = Left(cboDestOperacao.Text, 1)\r\n'
    b'\r\n'
    b'If sTipo = "1" And sFin = "1" And sDest = "1" Then\r\n'
    b'    cboNatureza.Text = "5102"\r\n'
    b'ElseIf sTipo = "1" And sFin = "1" And sDest = "2" Then\r\n'
    b'    cboNatureza.Text = "6102"\r\n'
    b'ElseIf sTipo = "1" And sFin = "1" And sDest = "3" Then\r\n'
    b'    cboNatureza.Text = "7102"\r\n'
    b'ElseIf sTipo = "1" And sFin = "4" And sDest = "1" Then\r\n'
    b'    cboNatureza.Text = "5202"\r\n'
    b'ElseIf sTipo = "1" And sFin = "4" And sDest = "2" Then\r\n'
    b'    cboNatureza.Text = "6202"\r\n'
    b'Else\r\n'
    b'    If cboNatureza.ListCount > 0 Then cboNatureza.Text = cboNatureza.List(0)\r\n'
    b'End If\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboDestOperacao_Change()\r\n'
)
c = data.count(old_change_anchor)
print(f'2. Inserir CarregarNaturezas+AjustarNaturezaPadrao: count={c}')
if c == 1: data = data.replace(old_change_anchor, new_subs)

# =============================================================================
# 3. cboDestOperacao_LostFocus: substituir linhas hardcoded por AjustarNaturezaPadrao
# =============================================================================
old_dest_nat = (
    b'If cboDestOperacao.Text = "1 - Opera\xe7\xe3o Interna" Then cboNatureza.Text = "5102"\r\n'
    b'If cboDestOperacao.Text = "2 - Opera\xe7\xe3o Interestadual" Then cboNatureza.Text = "6102"\r\n'
    b'AtualizarCFOPCSTItens\r\n'
    b'End Sub\r\n'
)
new_dest_nat = (
    b'AjustarNaturezaPadrao\r\n'
    b'AtualizarCFOPCSTItens\r\n'
    b'End Sub\r\n'
)
c = data.count(old_dest_nat)
print(f'3. cboDestOperacao_LostFocus natureza: count={c}')
if c == 1: data = data.replace(old_dest_nat, new_dest_nat)

# =============================================================================
# 4. cboFinalidade_Click: adicionar AjustarNaturezaPadrao
# =============================================================================
old_fin_click = (
    b'cboFinalidade_Click()\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    AplicarEstadoCheckboxes\r\n'
    b'    RecalcularItensNota\r\n'
    b'    CalcularICMSInterItensGERAL\r\n'
    b'End Sub\r\n'
)
new_fin_click = (
    b'cboFinalidade_Click()\r\n'
    b'    AjustarNaturezaPadrao\r\n'
    b'    AplicarVisibilidadeGridItens\r\n'
    b'    AplicarEstadoCheckboxes\r\n'
    b'    RecalcularItensNota\r\n'
    b'    CalcularICMSInterItensGERAL\r\n'
    b'End Sub\r\n'
)
c = data.count(old_fin_click)
print(f'4. cboFinalidade_Click: count={c}')
if c == 1: data = data.replace(old_fin_click, new_fin_click)

# =============================================================================
# 5. Adicionar cboTipoNota_LostFocus (dispara ao sair do combo apos escolha)
# =============================================================================
old_tipo_got = b'Private Sub cboTipoNota_GotFocus()\r\n'
new_tipo_got = (
    b'Private Sub cboTipoNota_LostFocus()\r\n'
    b'AjustarNaturezaPadrao\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub cboTipoNota_GotFocus()\r\n'
)
c = data.count(old_tipo_got)
print(f'5. cboTipoNota_LostFocus: count={c}')
if c == 1: data = data.replace(old_tipo_got, new_tipo_got)

# Normalizar CRLF
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
