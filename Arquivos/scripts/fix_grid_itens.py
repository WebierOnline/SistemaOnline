data = open('OnlineCommerce/Forms/NFe_Completa.frm', 'rb').read()
results = []

# ─────────────────────────────────────────────────────────────────────────────
# 1. Rewrite Exibir_Itens  (SELECT expandido + AplicarVisibilidadeGridItens)
# ─────────────────────────────────────────────────────────────────────────────
old = (
    b'Sub Exibir_Itens()\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'\r\n'
    b'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, pICMS, vICMS, ValorUnitarioComercializacao, QuantidadeComercial, valordesconto, ValorTotalBruto, IPIpIPI, IPIvIPI  FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'FormatarGridItensNota Tb\r\n'
    b'End Sub'
)
new = (
    b'Sub Exibir_Itens()\r\n'
    b'If txtCodNota.Text = "" Then Exit Sub\r\n'
    b'\r\n'
    b'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, " & _\r\n'
    b'       "ValorUnitarioComercializacao, QuantidadeComercial, ValorTotalBruto, " & _\r\n'
    b'       "ValorFrete, ValorSeguro, ValorOutros, ValorDesconto, " & _\r\n'
    b'       "vBC, pICMS, vICMS, pRedBC, " & _\r\n'
    b'       "vBCST, pICMSST, vICMSST, pMVAST, " & _\r\n'
    b'       "IPIpIPI, IPIvIPI, IPIcEnq " & _\r\n'
    b'       "FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'FormatarGridItensNota Tb\r\n'
    b'AplicarVisibilidadeGridItens\r\n'
    b'End Sub'
)
c = data.count(old)
results.append(f'1. Exibir_Itens: {c}')
if c == 1: data = data.replace(old, new)

# ─────────────────────────────────────────────────────────────────────────────
# 2. Rewrite FormatarGridItensNota  (27 colunas)
# ─────────────────────────────────────────────────────────────────────────────
old_fmt_start = b'Sub FormatarGridItensNota(rTabela As ADODB.Recordset)\r\n'
idx = data.find(old_fmt_start)
end_fmt = data.find(b'End Sub', idx) + 7
old_fmt = data[idx:end_fmt]

new_fmt = (
    b'Sub FormatarGridItensNota(rTabela As ADODB.Recordset)\r\n'
    b'   Dim i As Integer\r\n'
    b'   Dim j As Integer\r\n'
    b'\r\n'
    b'   With GridNotasItens\r\n'
    b'      .Visible = False\r\n'
    b'      .Redraw = False\r\n'
    b'\r\n'
    b'      .Clear\r\n'
    b'      .Cols = 27\r\n'
    b'      .rows = 2\r\n'
    b'\r\n'
    b'      \'Colunas fixas (sempre visiveis)\r\n'
    b'      .ColWidth(0)  = 200   \'indicador de linha\r\n'
    b'      .ColWidth(1)  = 400   \'No.\r\n'
    b'      .ColWidth(2)  = 1500  \'EAN\r\n'
    b'      .ColWidth(3)  = 0     \'COD. (oculto)\r\n'
    b'      .ColWidth(4)  = 3500  \'DESCRICAO\r\n'
    b'      .ColWidth(5)  = 450   \'UND\r\n'
    b'      .ColWidth(6)  = 900   \'NCM\r\n'
    b'      .ColWidth(7)  = 600   \'CFOP\r\n'
    b'      .ColWidth(8)  = 500   \'CST\r\n'
    b'      .ColWidth(9)  = 850   \'VALOR\r\n'
    b'      .ColWidth(10) = 850   \'QTDE\r\n'
    b'      .ColWidth(11) = 800   \'FRETE\r\n'
    b'      .ColWidth(12) = 800   \'SEGURO\r\n'
    b'      .ColWidth(13) = 800   \'OUTROS\r\n'
    b'      .ColWidth(14) = 800   \'DESC.\r\n'
    b'      .ColWidth(15) = 1050  \'TOTAL\r\n'
    b'      \'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)\r\n'
    b'      .ColWidth(16) = 0     \'BC ICMS\r\n'
    b'      .ColWidth(17) = 0     \'%ICMS\r\n'
    b'      .ColWidth(18) = 0     \'ICMS\r\n'
    b'      .ColWidth(19) = 0     \'%RED BC\r\n'
    b'      .ColWidth(20) = 0     \'BC ST\r\n'
    b'      .ColWidth(21) = 0     \'%ICMSST\r\n'
    b'      .ColWidth(22) = 0     \'ICMSST\r\n'
    b'      .ColWidth(23) = 0     \'MVA ST\r\n'
    b'      .ColWidth(24) = 0     \'%IPI\r\n'
    b'      .ColWidth(25) = 0     \'IPI\r\n'
    b'      .ColWidth(26) = 0     \'cEnq\r\n'
    b'\r\n'
    b'      .TextMatrix(0, 1)  = "No."\r\n'
    b'      .TextMatrix(0, 2)  = "EAN"\r\n'
    b'      .TextMatrix(0, 3)  = "C\xd3D."\r\n'
    b'      .TextMatrix(0, 4)  = "DESCRI\xc7\xc3O"\r\n'
    b'      .TextMatrix(0, 5)  = "UND"\r\n'
    b'      .TextMatrix(0, 6)  = "NCM"\r\n'
    b'      .TextMatrix(0, 7)  = "CFOP"\r\n'
    b'      .TextMatrix(0, 8)  = "CST"\r\n'
    b'      .TextMatrix(0, 9)  = "VALOR"\r\n'
    b'      .TextMatrix(0, 10) = "QTDE"\r\n'
    b'      .TextMatrix(0, 11) = "FRETE"\r\n'
    b'      .TextMatrix(0, 12) = "SEGURO"\r\n'
    b'      .TextMatrix(0, 13) = "OUTROS"\r\n'
    b'      .TextMatrix(0, 14) = "DESC."\r\n'
    b'      .TextMatrix(0, 15) = "TOTAL"\r\n'
    b'      .TextMatrix(0, 16) = "BC ICMS"\r\n'
    b'      .TextMatrix(0, 17) = "%ICMS"\r\n'
    b'      .TextMatrix(0, 18) = "ICMS"\r\n'
    b'      .TextMatrix(0, 19) = "%RED BC"\r\n'
    b'      .TextMatrix(0, 20) = "BC ST"\r\n'
    b'      .TextMatrix(0, 21) = "%ICMSST"\r\n'
    b'      .TextMatrix(0, 22) = "ICMSST"\r\n'
    b'      .TextMatrix(0, 23) = "MVA ST"\r\n'
    b'      .TextMatrix(0, 24) = "%IPI"\r\n'
    b'      .TextMatrix(0, 25) = "IPI"\r\n'
    b'      .TextMatrix(0, 26) = "cEnq"\r\n'
    b'\r\n'
    b'      \'Cabecalho em negrito e centralizado\r\n'
    b'      For i = 0 To .Cols - 1\r\n'
    b'         .Col = i: .Row = 0\r\n'
    b'         .CellFontBold = True\r\n'
    b'         .CellAlignment = flexAlignCenterCenter\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b'      \'Alinhamento: texto esquerda (0-8), numeros direita (9-26)\r\n'
    b'      For i = 0 To 8\r\n'
    b'         .ColAlignment(i) = 1\r\n'
    b'      Next i\r\n'
    b'      For i = 9 To 26\r\n'
    b'         .ColAlignment(i) = 6\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b'      i = 1\r\n'
    b'      If Not rTabela Is Nothing Then\r\n'
    b'         Do While Not rTabela.EOF\r\n'
    b'            .TextMatrix(.rows - 1, 1)  = Format(rTabela("ITEM"), "000")\r\n'
    b'            .TextMatrix(.rows - 1, 2)  = rTabela("EAN")\r\n'
    b'            .TextMatrix(.rows - 1, 3)  = Format(rTabela("CodigoProduto"), "00000")\r\n'
    b'            .TextMatrix(.rows - 1, 4)  = rTabela("NomeProduto")\r\n'
    b'            .TextMatrix(.rows - 1, 5)  = rTabela("UnidadeComercial")\r\n'
    b'            .TextMatrix(.rows - 1, 6)  = rTabela("NCM")\r\n'
    b'            .TextMatrix(.rows - 1, 7)  = rTabela("CFOP")\r\n'
    b'            .TextMatrix(.rows - 1, 8)  = rTabela("CST")\r\n'
    b'            .TextMatrix(.rows - 1, 9)  = FormatNumber(rTabela("ValorUnitarioComercializacao"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 10) = Format(rTabela("QuantidadeComercial"), ocPESO)\r\n'
    b'            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela("ValorFrete"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela("ValorSeguro"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela("ValorOutros"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela("ValorDesconto"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 15) = FormatNumber(rTabela("ValorTotalBruto"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela("vBC"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela("pICMS"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela("vICMS"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela("pRedBC"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela("vBCST"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela("pICMSST"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela("vICMSST"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela("pMVAST"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela("IPIpIPI"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela("IPIvIPI"), 2)\r\n'
    b'            .TextMatrix(.rows - 1, 26) = rTabela("IPIcEnq")\r\n'
    b'\r\n'
    b'            rTabela.MoveNext\r\n'
    b'            .rows = .rows + 1\r\n'
    b'            i = i + 1\r\n'
    b'         Loop\r\n'
    b'      End If\r\n'
    b'\r\n'
    b'      .rows = .rows - 1\r\n'
    b'\r\n'
    b'      \'EAN em negrito\r\n'
    b'      For i = 1 To .rows - 1\r\n'
    b'         .Row = i: .Col = 2: .CellFontBold = True\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b'      \'COD. em destaque\r\n'
    b'      For i = 1 To .rows - 1\r\n'
    b'         .Row = i: .Col = 3\r\n'
    b'         .CellForeColor = &HC0&\r\n'
    b'         .CellFontBold = True\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b'      \'TOTAL em destaque\r\n'
    b'      For i = 1 To .rows - 1\r\n'
    b'         .Row = i: .Col = 15\r\n'
    b'         .CellForeColor = &HC0&\r\n'
    b'         .CellFontBold = True\r\n'
    b'      Next i\r\n'
    b'\r\n'
    b'      GridNotasItens.Col = 0\r\n'
    b'      .Visible = True\r\n'
    b'      .Redraw = True\r\n'
    b'   End With\r\n'
    b'End Sub'
)
c = data.count(old_fmt)
results.append(f'2. FormatarGridItensNota: {c}')
if c == 1: data = data.replace(old_fmt, new_fmt)

# ─────────────────────────────────────────────────────────────────────────────
# 3. Insert AplicarVisibilidadeGridItens + checkbox handlers after End Sub of FormatarGridItensNota
# ─────────────────────────────────────────────────────────────────────────────
anchor = (
    b'      GridNotasItens.Col = 0\r\n'
    b'      .Visible = True\r\n'
    b'      .Redraw = True\r\n'
    b'   End With\r\n'
    b'End Sub'
)
insert_after = anchor + (
    b'\r\n'
    b'\r\n'
    b'Private Sub AplicarVisibilidadeGridItens()\r\n'
    b'   \'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)\r\n'
    b'   Dim bICMS As Boolean\r\n'
    b'   bICMS = (Left(cboFinalidade.Text, 1) = "4")\r\n'
    b'   GridNotasItens.ColWidth(16) = IIf(bICMS, 850, 0)  \'BC ICMS\r\n'
    b'   GridNotasItens.ColWidth(17) = IIf(bICMS, 700, 0)  \'%ICMS\r\n'
    b'   GridNotasItens.ColWidth(18) = IIf(bICMS, 850, 0)  \'ICMS\r\n'
    b'\r\n'
    b'   \'%RedBC: chkpRedBC\r\n'
    b'   GridNotasItens.ColWidth(19) = IIf(chkpRedBC.Value = 1, 700, 0)\r\n'
    b'\r\n'
    b'   \'Grupo ICMSST: chkICMSST\r\n'
    b'   Dim bST As Boolean\r\n'
    b'   bST = (chkICMSST.Value = 1)\r\n'
    b'   GridNotasItens.ColWidth(20) = IIf(bST, 850, 0)  \'BC ST\r\n'
    b'   GridNotasItens.ColWidth(21) = IIf(bST, 700, 0)  \'%ICMSST\r\n'
    b'   GridNotasItens.ColWidth(22) = IIf(bST, 850, 0)  \'ICMSST\r\n'
    b'   GridNotasItens.ColWidth(23) = IIf(bST, 700, 0)  \'MVA ST\r\n'
    b'\r\n'
    b'   \'Grupo IPI: chkIPI\r\n'
    b'   Dim bIPI As Boolean\r\n'
    b'   bIPI = (chkIPI.Value = 1)\r\n'
    b'   GridNotasItens.ColWidth(24) = IIf(bIPI, 700, 0)  \'%IPI\r\n'
    b'   GridNotasItens.ColWidth(25) = IIf(bIPI, 850, 0)  \'IPI\r\n'
    b'   GridNotasItens.ColWidth(26) = IIf(bIPI, 700, 0)  \'cEnq\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub chkIPI_Click()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub chkpRedBC_Click()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub\r\n'
    b'\r\n'
    b'Private Sub chkICMSST_Click()\r\n'
    b'   AplicarVisibilidadeGridItens\r\n'
    b'End Sub'
)
c = data.count(anchor)
results.append(f'3. Inserir AplicarVisibilidade+checkboxes: {c}')
if c == 1: data = data.replace(anchor, insert_after)

# ─────────────────────────────────────────────────────────────────────────────
# 4. Callers: substituir SELECT antigo + RsOpen + FormatarGridItensNota por Exibir_Itens
#    (TransformarPedidoemNFE, cmdDuplicar_Click, cmdDuplicarCFOP_Click)
# ─────────────────────────────────────────────────────────────────────────────
old_select = (
    b'sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, pICMS, vICMS, ValorUnitarioComercializacao, QuantidadeComercial, valordesconto, ValorTotalBruto, IPIpIPI, IPIvIPI FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'RsOpen Tb, sSQL\r\n'
    b'\r\n'
    b'FormatarGridItensNota Tb'
)
c = data.count(old_select)
results.append(f'4. Callers SELECT+RsOpen+Format: {c}')
if c >= 1: data = data.replace(old_select, b'Exibir_Itens')

# ─────────────────────────────────────────────────────────────────────────────
# 5. cmdRemoverItem_Click: mesma substituicao mas com indentacao
# ─────────────────────────────────────────────────────────────────────────────
old_remover = (
    b'    sSQL = "SELECT ITEM, EAN, CodigoProduto, NomeProduto, UnidadeComercial, NCM, CFOP, CST, pICMS, vICMS, ValorUnitarioComercializacao, QuantidadeComercial, valordesconto, ValorTotalBruto, IPIpIPI, IPIvIPI FROM NotaFiscalItens WHERE CodigoNota = " & Val(txtCodNota.Text)\r\n'
    b'    RsOpen Tb, sSQL\r\n'
    b'    \r\n'
    b'    FormatarGridItensNota Tb\r\n'
    b'    \r\n'
    b'\'    lblValorNota.Caption = Format(Tb("vTotal"), ocMONEY)\r\n'
    b'    \r\n'
)
new_remover = b'    Exibir_Itens\r\n'
c = data.count(old_remover)
results.append(f'5. cmdRemoverItem_Click: {c}')
if c == 1: data = data.replace(old_remover, new_remover)

# ─────────────────────────────────────────────────────────────────────────────
# 6. Mostrar_ItensNota: LimparGridItensNota + DoEvents + FormatarGridItensNota
# ─────────────────────────────────────────────────────────────────────────────
old_mostrar = b'LimparGridItensNota\r\nDoEvents\r\nFormatarGridItensNota Tb\r\n'
new_mostrar = b'Exibir_Itens\r\n'
c = data.count(old_mostrar)
results.append(f'6. Mostrar_ItensNota: {c}')
if c == 1: data = data.replace(old_mostrar, new_mostrar)

# ─────────────────────────────────────────────────────────────────────────────
# Salvar
# ─────────────────────────────────────────────────────────────────────────────
data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/NFe_Completa.frm', 'wb').write(data)

for r in results: print(r)
print('Salvo. Tamanho:', len(data))
