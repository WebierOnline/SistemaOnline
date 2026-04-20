path = 'C:/projeto/OnlineCommerce/Forms/Entrada_Estoque.Frm'
data = open(path, 'rb').read().decode('windows-1252')
lines = data.split('\r\n')

# ============================================================
# R1: Substitui bloco Frame1 (linhas 618-864) por Frame1+grdTotaisNota
# ============================================================

# Localiza Frame1 pelo strip (exato)
start_line = -1
for i, line in enumerate(lines):
    if line.strip() == 'Begin VB.Frame Frame1':
        start_line = i
        break

assert start_line >= 0, "Frame1 nao encontrado"

# Encontra End correspondente pelo contador de profundidade
depth = 0
end_line = -1
for i in range(start_line, len(lines)):
    s = lines[i].strip()
    if (s.startswith('Begin ') or s.startswith('Begin MSFlexGridLib')) and not s.startswith('BeginProperty'):
        depth += 1
    elif s == 'End':
        depth -= 1
        if depth == 0:
            end_line = i
            break

assert end_line >= 0, "End do Frame1 nao encontrado"
print('Frame1 block: lines', start_line + 1, 'to', end_line + 1)

new_frame1_lines = [
    "         Begin VB.Frame Frame1 ",
    "            Caption         =   \"Totais da Nota\"",
    "            BeginProperty Font ",
    "               Name            =   \"MS Sans Serif\"",
    "               Size            =   8.25",
    "               Charset         =   0",
    "               Weight          =   700",
    "               Underline       =   0   'False",
    "               Italic          =   0   'False",
    "               Strikethrough   =   0   'False",
    "            EndProperty",
    "            Height          =   2715",
    "            Left            =   -74880",
    "            TabIndex        =   162",
    "            Top             =   360",
    "            Width           =   15375",
    "            Begin MSFlexGridLib.MSFlexGrid grdTotaisNota ",
    "               Height          =   2355",
    "               Left            =   60",
    "               TabIndex        =   226",
    "               Top             =   300",
    "               Width           =   15255",
    "               _ExtentX        =   26906",
    "               _ExtentY        =   4154",
    "               _Version        =   393216",
    "               SelectionMode   =   1",
    "               Appearance      =   0",
    "            End",
    "         End",
]

lines[start_line:end_line + 1] = new_frame1_lines
data2 = '\r\n'.join(lines) + '\r\n'
print('r1 (Frame1 substituido): ok')

# ============================================================
# R2: Remove TbEntrada("ValorFrete") = ... na rotina Salvar_Dados
# ============================================================
old2 = '    TbEntrada("ValorFrete") = IIf(Vazio(txtValorFrete), 0, CDbl(Format(txtValorFrete, "##0.00")))\r\n'
print('r2 found:', data2.count(old2))
data2 = data2.replace(old2, '', 1)

# ============================================================
# R3: Substitui os 12 assignments txt... = Format(TbEntrada(...))
#     por CarregarTotaisNota TbEntrada
# ============================================================
old3 = (
    '    txtValorSeguro = Format(TbEntrada("ValorSeguro"), "##,##0.00")\r\n'
    '    txtValorOutrasDespesas = Format(TbEntrada("ValorOutrasDespesas"), "##,##0.00")\r\n'
    '    txtValorFrete = Format(TbEntrada("ValorFrete"), "##,##0.00")\r\n'
    '    txtBaseICMS = Format(TbEntrada("BaseICMS"), "##,##0.00")\r\n'
    '    txtValorICMS = Format(TbEntrada("ValorICMS"), "##,##0.00")\r\n'
    '    txtBaseICMSST = Format(TbEntrada("BaseICMSST"), "##,##0.00")\r\n'
    '    txtValorICMSST = Format(TbEntrada("ValorICMSST"), "##,##0.00")\r\n'
    '    txtValorIPI = Format(TbEntrada("ValorIPI"), "##,##0.00")\r\n'
    '    txtValorPIS = Format(TbEntrada("ValorPIS"), "##,##0.00")\r\n'
    '    txtValorCOFINS = Format(TbEntrada("ValorCOFINS"), "##,##0.00")\r\n'
    '    txtValorDesconto = Format(TbEntrada("ValorDesconto"), "##,##0.00")\r\n'
    '    txtTotaldosProdutos = Format(TbEntrada("ValorProdutos"), "##,##0.00")\r\n'
)
new3 = '    CarregarTotaisNota TbEntrada\r\n'
print('r3 found:', data2.count(old3))
data2 = data2.replace(old3, new3, 1)

# ============================================================
# R4: Remove txtValorFrete = Format(0, ...) na rotina de limpeza
# ============================================================
old4 = '    txtValorFrete = Format(0, "##,##0.00")\r\n'
print('r4 found:', data2.count(old4))
data2 = data2.replace(old4, '', 1)

# ============================================================
# R5: Acrescenta subs auxiliares + CarregarTotaisNota no final
# ============================================================
new_subs = r"""
Private Sub grdTotHeader(g As MSFlexGrid, sTitle As String)
   g.Rows = g.Rows + 1
   Dim r As Long
   r = g.Rows - 1
   g.TextMatrix(r, 0) = ""
   g.TextMatrix(r, 1) = sTitle
   g.TextMatrix(r, 2) = ""
   g.Row = r: g.Col = 0: g.CellBackColor = &HC0C0C0
   g.Row = r: g.Col = 1: g.CellBackColor = &HC0C0C0: g.CellFontBold = True
   g.Row = r: g.Col = 2: g.CellBackColor = &HC0C0C0
End Sub

Private Sub grdTotAdd(g As MSFlexGrid, sSec As String, sDesc As String, sVal As String)
   g.Rows = g.Rows + 1
   Dim r As Long
   r = g.Rows - 1
   g.TextMatrix(r, 0) = sSec
   g.TextMatrix(r, 1) = sDesc
   g.TextMatrix(r, 2) = sVal
   g.Row = r: g.Col = 2: g.CellAlignment = flexAlignRightCenter
End Sub

Private Sub grdTotTotal(g As MSFlexGrid, sDesc As String, sVal As String)
   g.Rows = g.Rows + 1
   Dim r As Long
   r = g.Rows - 1
   g.TextMatrix(r, 0) = "TOTAL"
   g.TextMatrix(r, 1) = sDesc
   g.TextMatrix(r, 2) = sVal
   Dim c As Integer
   For c = 0 To 2
      g.Row = r: g.Col = c
      g.CellFontBold = True
      g.CellBackColor = &HFFFFC0
   Next c
   g.Row = r: g.Col = 2: g.CellAlignment = flexAlignRightCenter
End Sub

Private Sub CarregarTotaisNota(TbEnt As ADODB.Recordset)
   If TbEnt Is Nothing Then Exit Sub
   Dim F As String
   F = "##,##0.00"
   With grdTotaisNota
      .Redraw = False
      .Rows = 1
      .Cols = 3
      .FixedRows = 1
      .FixedCols = 0
      .ColWidth(0) = 2400
      .ColWidth(1) = 9000
      .ColWidth(2) = 3855
      .TextMatrix(0, 0) = "Secao"
      .TextMatrix(0, 1) = "Descricao"
      .TextMatrix(0, 2) = "Valor"
   End With
   Call grdTotHeader(grdTotaisNota, "ICMS")
   Call grdTotAdd(grdTotaisNota, "ICMS", "Base de Calc. ICMS (vBC)", Format(TbEnt("BaseICMS"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS", "Valor do ICMS", Format(TbEnt("ValorICMS"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS", "Valor do FCP", Format(TbEnt("ValorFCP"), F))
   Call grdTotHeader(grdTotaisNota, "ICMS-ST")
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Base de Calc. ICMS-ST", Format(TbEnt("BaseICMSST"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor do ICMS-ST", Format(TbEnt("ValorICMSST"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor do FCP-ST", Format(TbEnt("ValorFCPST"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Valor FCP-ST Retido", Format(TbEnt("ValorFCPSTRet"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "Base ICMS-ST Retido (Total)", Format(TbEnt("vBCSTRetTotal"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "ICMS-ST Retido (Total)", Format(TbEnt("vICMSSTRetTotal"), F))
   Call grdTotAdd(grdTotaisNota, "ICMS-ST", "ICMS Substituto (Total)", Format(TbEnt("vICMSSubstitutoTotal"), F))
   Call grdTotHeader(grdTotaisNota, "Produtos e Servicos")
   Call grdTotAdd(grdTotaisNota, "Produtos", "Total dos Produtos", Format(TbEnt("ValorProdutos"), F))
   Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Frete", Format(TbEnt("ValorFrete"), F))
   Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Seguro", Format(TbEnt("ValorSeguro"), F))
   Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do Desconto", Format(TbEnt("ValorDesconto"), F))
   Call grdTotAdd(grdTotaisNota, "Produtos", "Outras Despesas", Format(TbEnt("ValorOutrasDespesas"), F))
   Call grdTotAdd(grdTotaisNota, "Produtos", "Valor do II (Importacao)", Format(TbEnt("ValorImportacao"), F))
   Call grdTotHeader(grdTotaisNota, "IPI / PIS / COFINS")
   Call grdTotAdd(grdTotaisNota, "IPI", "Valor do IPI", Format(TbEnt("ValorIPI"), F))
   Call grdTotAdd(grdTotaisNota, "IPI", "Valor IPI Devolvido", Format(TbEnt("ValorIPIDevol"), F))
   Call grdTotAdd(grdTotaisNota, "PIS", "Valor do PIS", Format(TbEnt("ValorPIS"), F))
   Call grdTotAdd(grdTotaisNota, "COFINS", "Valor do COFINS", Format(TbEnt("ValorCOFINS"), F))
   Call grdTotHeader(grdTotaisNota, "DIFAL")
   Call grdTotAdd(grdTotaisNota, "DIFAL", "FCP UF Destinataria", Format(TbEnt("vFCPUFDest"), F))
   Call grdTotAdd(grdTotaisNota, "DIFAL", "ICMS UF Destinataria", Format(TbEnt("vICMSUFDest"), F))
   Call grdTotAdd(grdTotaisNota, "DIFAL", "ICMS UF Remetente", Format(TbEnt("vICMSUFRemet"), F))
   Call grdTotTotal(grdTotaisNota, "VALOR TOTAL DA NOTA", Format(TbEnt("ValorNota"), F))
   Call grdTotHeader(grdTotaisNota, "IBS / CBS (Reforma Tributaria)")
   Call grdTotAdd(grdTotaisNota, "IBS/CBS", "Base de Calc. IBS/CBS (vBC)", Format(TbEnt("TotIBSCBSvBC"), F))
   Call grdTotAdd(grdTotaisNota, "IBS", "IBS Estado (UF)", Format(TbEnt("TotIBSUFvIBS"), F))
   Call grdTotAdd(grdTotaisNota, "IBS", "IBS Municipio", Format(TbEnt("TotIBSMunvIBS"), F))
   Call grdTotAdd(grdTotaisNota, "IBS", "Total IBS", Format(TbEnt("TotIBSvIBS"), F))
   Call grdTotAdd(grdTotaisNota, "IBS", "Credito Presumido IBS", Format(TbEnt("TotIBSvCredPres"), F))
   Call grdTotAdd(grdTotaisNota, "CBS", "Total CBS", Format(TbEnt("TotCBSvCBS"), F))
   Call grdTotAdd(grdTotaisNota, "CBS", "Credito Presumido CBS", Format(TbEnt("TotCBSvCredPres"), F))
   Call grdTotAdd(grdTotaisNota, "IBS/CBS", "Valor NF com IBS/CBS (vNFTot)", Format(TbEnt("vNFTot"), F))
   grdTotaisNota.Redraw = True
End Sub
"""

# Normaliza quebras no novo_subs para CRLF
new_subs = new_subs.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')

# Insere antes do ultimo End Sub do arquivo
last_end_sub = data2.rfind('\r\nEnd Sub\r\n')
assert last_end_sub >= 0, "End Sub final nao encontrado"
insert_at = last_end_sub + len('\r\nEnd Sub\r\n')
data2 = data2[:insert_at] + new_subs
print('r5 (CarregarTotaisNota + helpers): appended')

# ============================================================
# Grava
# ============================================================
print('changed:', data2 != data)
raw = data2.encode('windows-1252')
raw = raw.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open(path, 'wb').write(raw)
print('ok')
