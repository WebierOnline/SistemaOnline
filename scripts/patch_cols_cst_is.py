"""
patch_cols_cst_is.py
Insere as colunas CST IS (col 11) e CLASS IS (col 12) no GridNotasItens.
Todas as colunas antigas >= 11 deslocam +2 (totalizando 34 colunas, 0-33).
Adiciona chkReformaIS_Click. Atualiza Exibir_Itens, AplicarVisibilidadeGridItens
e todos os Case handlers em txtEdit_LostFocus.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cols_is"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ── P1: .Cols = 32 → 34 ──────────────────────────────────────────────────────
patches.append((
    b"      .Cols = 32\r\n",
    b"      .Cols = 34\r\n"
))

# ── P2: ColWidth block cols 10-31 → 10-33 ────────────────────────────────────
patches.append((
    b"      .ColWidth(10) = 0     'cClassTrib (chkReforma)\r\n"
    b"      .ColWidth(11) = 0     'V. IBS\r\n"
    b"      .ColWidth(12) = 0     'V. CBS\r\n"
    b"      .ColWidth(13) = 0     'V. IS\r\n"
    b"      .ColWidth(14) = 850   'VALOR\r\n"
    b"      .ColWidth(15) = 850   'QTDE\r\n"
    b"      .ColWidth(16) = 800   'FRETE\r\n"
    b"      .ColWidth(17) = 0     'SEGURO (chkSeguro) - oculto por padrao\r\n"
    b"      .ColWidth(18) = 0     'OUTROS (chkOutros) - oculto por padrao\r\n"
    b"      .ColWidth(19) = 800   'DESC.\r\n"
    b"      .ColWidth(20) = 1050  'TOTAL\r\n"
    b"      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)\r\n"
    b"      .ColWidth(21) = 0     'BC ICMS\r\n"
    b"      .ColWidth(22) = 0     '%ICMS\r\n"
    b"      .ColWidth(23) = 0     'ICMS\r\n"
    b"      .ColWidth(24) = 0     '%RED BC\r\n"
    b"      .ColWidth(25) = 0     'BC ST\r\n"
    b"      .ColWidth(26) = 0     '%ICMSST\r\n"
    b"      .ColWidth(27) = 0     'ICMSST\r\n"
    b"      .ColWidth(28) = 0     'MVA ST\r\n"
    b"      .ColWidth(29) = 0     'CST IPI\r\n"
    b"      .ColWidth(30) = 0     '%IPI\r\n"
    b"      .ColWidth(31) = 0     'IPI\r\n",

    b"      .ColWidth(10) = 0     'cClassTrib (chkReforma)\r\n"
    b"      .ColWidth(11) = 0     'CST IS (chkReformaIS)\r\n"
    b"      .ColWidth(12) = 0     'CLASS IS (chkReformaIS)\r\n"
    b"      .ColWidth(13) = 0     'V. IBS\r\n"
    b"      .ColWidth(14) = 0     'V. CBS\r\n"
    b"      .ColWidth(15) = 0     'V. IS\r\n"
    b"      .ColWidth(16) = 850   'VALOR\r\n"
    b"      .ColWidth(17) = 850   'QTDE\r\n"
    b"      .ColWidth(18) = 800   'FRETE\r\n"
    b"      .ColWidth(19) = 0     'SEGURO (chkSeguro) - oculto por padrao\r\n"
    b"      .ColWidth(20) = 0     'OUTROS (chkOutros) - oculto por padrao\r\n"
    b"      .ColWidth(21) = 800   'DESC.\r\n"
    b"      .ColWidth(22) = 1050  'TOTAL\r\n"
    b"      'Colunas condicionais (largura definida por AplicarVisibilidadeGridItens)\r\n"
    b"      .ColWidth(23) = 0     'BC ICMS\r\n"
    b"      .ColWidth(24) = 0     '%ICMS\r\n"
    b"      .ColWidth(25) = 0     'ICMS\r\n"
    b"      .ColWidth(26) = 0     '%RED BC\r\n"
    b"      .ColWidth(27) = 0     'BC ST\r\n"
    b"      .ColWidth(28) = 0     '%ICMSST\r\n"
    b"      .ColWidth(29) = 0     'ICMSST\r\n"
    b"      .ColWidth(30) = 0     'MVA ST\r\n"
    b"      .ColWidth(31) = 0     'CST IPI\r\n"
    b"      .ColWidth(32) = 0     '%IPI\r\n"
    b"      .ColWidth(33) = 0     'IPI\r\n"
))

# ── P3: TextMatrix headers cols 11-31 → 11-33 ────────────────────────────────
patches.append((
    b"      .TextMatrix(0, 11) = \"V. IBS\"\r\n"
    b"      .TextMatrix(0, 12) = \"V. CBS\"\r\n"
    b"      .TextMatrix(0, 13) = \"V. IS\"\r\n"
    b"      .TextMatrix(0, 14) = \"VALOR\"\r\n"
    b"      .TextMatrix(0, 15) = \"QTDE\"\r\n"
    b"      .TextMatrix(0, 16) = \"FRETE\"\r\n"
    b"      .TextMatrix(0, 17) = \"SEGURO\"\r\n"
    b"      .TextMatrix(0, 18) = \"OUTROS\"\r\n"
    b"      .TextMatrix(0, 19) = \"DESC.\"\r\n"
    b"      .TextMatrix(0, 20) = \"TOTAL\"\r\n"
    b"      .TextMatrix(0, 21) = \"BC ICMS\"\r\n"
    b"      .TextMatrix(0, 22) = \"%ICMS\"\r\n"
    b"      .TextMatrix(0, 23) = \"ICMS\"\r\n"
    b"      .TextMatrix(0, 24) = \"%RED BC\"\r\n"
    b"      .TextMatrix(0, 25) = \"BC ST\"\r\n"
    b"      .TextMatrix(0, 26) = \"%ICMSST\"\r\n"
    b"      .TextMatrix(0, 27) = \"ICMSST\"\r\n"
    b"      .TextMatrix(0, 28) = \"MVA ST\"\r\n"
    b"      .TextMatrix(0, 29) = \"CST IPI\"\r\n"
    b"      .TextMatrix(0, 30) = \"%IPI\"\r\n"
    b"      .TextMatrix(0, 31) = \"IPI\"\r\n",

    b"      .TextMatrix(0, 11) = \"CST IS\"\r\n"
    b"      .TextMatrix(0, 12) = \"CLASS IS\"\r\n"
    b"      .TextMatrix(0, 13) = \"V. IBS\"\r\n"
    b"      .TextMatrix(0, 14) = \"V. CBS\"\r\n"
    b"      .TextMatrix(0, 15) = \"V. IS\"\r\n"
    b"      .TextMatrix(0, 16) = \"VALOR\"\r\n"
    b"      .TextMatrix(0, 17) = \"QTDE\"\r\n"
    b"      .TextMatrix(0, 18) = \"FRETE\"\r\n"
    b"      .TextMatrix(0, 19) = \"SEGURO\"\r\n"
    b"      .TextMatrix(0, 20) = \"OUTROS\"\r\n"
    b"      .TextMatrix(0, 21) = \"DESC.\"\r\n"
    b"      .TextMatrix(0, 22) = \"TOTAL\"\r\n"
    b"      .TextMatrix(0, 23) = \"BC ICMS\"\r\n"
    b"      .TextMatrix(0, 24) = \"%ICMS\"\r\n"
    b"      .TextMatrix(0, 25) = \"ICMS\"\r\n"
    b"      .TextMatrix(0, 26) = \"%RED BC\"\r\n"
    b"      .TextMatrix(0, 27) = \"BC ST\"\r\n"
    b"      .TextMatrix(0, 28) = \"%ICMSST\"\r\n"
    b"      .TextMatrix(0, 29) = \"ICMSST\"\r\n"
    b"      .TextMatrix(0, 30) = \"MVA ST\"\r\n"
    b"      .TextMatrix(0, 31) = \"CST IPI\"\r\n"
    b"      .TextMatrix(0, 32) = \"%IPI\"\r\n"
    b"      .TextMatrix(0, 33) = \"IPI\"\r\n"
))

# ── P4: Alinhamento numérico 9 To 31 → 9 To 33 ───────────────────────────────
patches.append((
    b"      For i = 9 To 31\r\n"
    b"         .ColAlignment(i) = 6\r\n"
    b"      Next i\r\n",

    b"      For i = 9 To 33\r\n"
    b"         .ColAlignment(i) = 6\r\n"
    b"      Next i\r\n"
))

# ── P5: Data fill cols 11-31 → novo 11/12 + 13-33 ────────────────────────────
patches.append((
    b"            .TextMatrix(.rows - 1, 11) = FormatNumber(rTabela(\"IBS_vIBS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 12) = FormatNumber(rTabela(\"CBS_vCBS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela(\"IS_vIS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela(\"ValorUnitarioComercializacao\"), 2)\r\n"
    b"            If rTabela(\"UnidadeComercial\") = \"KG\" Or rTabela(\"UnidadeComercial\") = \"GR\" Or rTabela(\"UnidadeComercial\") = \"MG\" Then\r\n"
    b"                .TextMatrix(.rows - 1, 15) = Format(rTabela(\"QuantidadeComercial\"), ocPESO)\r\n"
    b"            Else\r\n"
    b"                .TextMatrix(.rows - 1, 15) = Format(rTabela(\"QuantidadeComercial\"), \"###,###,##0\")\r\n"
    b"            End If\r\n"
    b"            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela(\"ValorFrete\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 17) = FormatNumber(rTabela(\"ValorSeguro\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela(\"ValorOutros\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela(\"ValorDesconto\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela(\"ValorTotalBruto\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela(\"vBC\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela(\"pICMS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela(\"vICMS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela(\"pRedBC\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela(\"vBCST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela(\"pICMSST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 27) = FormatNumber(rTabela(\"vICMSST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 28) = FormatNumber(rTabela(\"pMVAST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 29) = rTabela(\"IPICST\")\r\n"
    b"            .TextMatrix(.rows - 1, 30) = FormatNumber(rTabela(\"IPIpIPI\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 31) = FormatNumber(rTabela(\"IPIvIPI\"), 2)\r\n",

    b"            .TextMatrix(.rows - 1, 11) = ValidateNull(rTabela(\"IS_CST\"))\r\n"
    b"            .TextMatrix(.rows - 1, 12) = ValidateNull(rTabela(\"cClassTrib_IS\"))\r\n"
    b"            .TextMatrix(.rows - 1, 13) = FormatNumber(rTabela(\"IBS_vIBS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 14) = FormatNumber(rTabela(\"CBS_vCBS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 15) = FormatNumber(rTabela(\"IS_vIS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 16) = FormatNumber(rTabela(\"ValorUnitarioComercializacao\"), 2)\r\n"
    b"            If rTabela(\"UnidadeComercial\") = \"KG\" Or rTabela(\"UnidadeComercial\") = \"GR\" Or rTabela(\"UnidadeComercial\") = \"MG\" Then\r\n"
    b"                .TextMatrix(.rows - 1, 17) = Format(rTabela(\"QuantidadeComercial\"), ocPESO)\r\n"
    b"            Else\r\n"
    b"                .TextMatrix(.rows - 1, 17) = Format(rTabela(\"QuantidadeComercial\"), \"###,###,##0\")\r\n"
    b"            End If\r\n"
    b"            .TextMatrix(.rows - 1, 18) = FormatNumber(rTabela(\"ValorFrete\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 19) = FormatNumber(rTabela(\"ValorSeguro\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 20) = FormatNumber(rTabela(\"ValorOutros\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 21) = FormatNumber(rTabela(\"ValorDesconto\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 22) = FormatNumber(rTabela(\"ValorTotalBruto\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 23) = FormatNumber(rTabela(\"vBC\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 24) = FormatNumber(rTabela(\"pICMS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 25) = FormatNumber(rTabela(\"vICMS\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 26) = FormatNumber(rTabela(\"pRedBC\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 27) = FormatNumber(rTabela(\"vBCST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 28) = FormatNumber(rTabela(\"pICMSST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 29) = FormatNumber(rTabela(\"vICMSST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 30) = FormatNumber(rTabela(\"pMVAST\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 31) = rTabela(\"IPICST\")\r\n"
    b"            .TextMatrix(.rows - 1, 32) = FormatNumber(rTabela(\"IPIpIPI\"), 2)\r\n"
    b"            .TextMatrix(.rows - 1, 33) = FormatNumber(rTabela(\"IPIvIPI\"), 2)\r\n"
))

# ── P6: DESC. highlight .Col = 19 → 21 ───────────────────────────────────────
patches.append((
    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i: .Col = 19\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next i\r\n"
    b"\r\n"
    b"      'Colunas edit\xe1veis em amarelo claro\r\n",

    b"      For i = 1 To .rows - 1\r\n"
    b"         .Row = i: .Col = 21\r\n"
    b"         .CellForeColor = &HC0&\r\n"
    b"         .CellFontBold = True\r\n"
    b"      Next i\r\n"
    b"\r\n"
    b"      'Colunas edit\xe1veis em amarelo claro\r\n"
))

# ── P7: colEdit array ─────────────────────────────────────────────────────────
patches.append((
    b"      For Each colEdit In Array(2, 5, 6, 7, 8, 22, 24, 26, 28, 29, 30)",
    b"      For Each colEdit In Array(2, 5, 6, 7, 8, 24, 26, 28, 30, 31, 32)"
))

# ── P8: colRef array ──────────────────────────────────────────────────────────
patches.append((
    b"      For Each colRef In Array(9, 10, 11, 12, 13)",
    b"      For Each colRef In Array(9, 10, 11, 12, 13, 14, 15)"
))

# ── P9: AplicarVisibilidadeGridItens — substituição completa do corpo ─────────
patches.append((
    b"Sub AplicarVisibilidadeGridItens()\r\n"
    b"   If GridNotasItens.Cols < 32 Then Exit Sub\r\n"
    b"   'Reforma tributaria: chkReforma\r\n"
    b"   Dim bReforma As Boolean\r\n"
    b"   bReforma = (chkReforma.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(9) = IIf(bReforma, 700, 0)     'CST IBS/CBS\r\n"
    b"   GridNotasItens.ColWidth(10) = IIf(bReforma, 1200, 0)   'cClassTrib\r\n"
    b"   GridNotasItens.ColWidth(11) = IIf(bReforma, 850, 0)    'V. IBS\r\n"
    b"   GridNotasItens.ColWidth(12) = IIf(bReforma, 850, 0)    'V. CBS\r\n"
    b"   GridNotasItens.ColWidth(13) = IIf(bReforma, 850, 0)    'V. IS\r\n"
    b"\r\n"
    b"   'Seguro: chkSeguro\r\n"
    b"   GridNotasItens.ColWidth(17) = IIf(chkSeguro.Value = 1, 900, 0)\r\n"
    b"   'Outros: chkOutros\r\n"
    b"   GridNotasItens.ColWidth(18) = IIf(chkOutros.Value = 1, 900, 0)\r\n"
    b"\r\n"
    b"   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)\r\n"
    b"   Dim bICMS As Boolean\r\n"
    b"   bICMS = (Left(cboFinalidade.Text, 1) = \"4\")\r\n"
    b"   GridNotasItens.ColWidth(21) = IIf(bICMS, 850, 0)  'BC ICMS\r\n"
    b"   GridNotasItens.ColWidth(22) = IIf(bICMS, 850, 0)  '%ICMS\r\n"
    b"   GridNotasItens.ColWidth(23) = IIf(bICMS, 850, 0)  'ICMS\r\n"
    b"\r\n"
    b"   '%RedBC: chkpRedBC\r\n"
    b"   GridNotasItens.ColWidth(24) = IIf(chkpRedBC.Value = 1, 700, 0)\r\n"
    b"\r\n"
    b"   'Grupo ICMSST: chkICMSST\r\n"
    b"   Dim bST As Boolean\r\n"
    b"   bST = (chkICMSST.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(25) = IIf(bST, 850, 0)  'BC ST\r\n"
    b"   GridNotasItens.ColWidth(26) = IIf(bST, 900, 0)  '%ICMSST\r\n"
    b"   GridNotasItens.ColWidth(27) = IIf(bST, 850, 0)  'ICMSST\r\n"
    b"   GridNotasItens.ColWidth(28) = IIf(bST, 850, 0)  'MVA ST\r\n"
    b"\r\n"
    b"   'Grupo IPI: chkIPI\r\n"
    b"   Dim bIPI As Boolean\r\n"
    b"   bIPI = (chkIPI.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(29) = IIf(bIPI, 850, 0)  'CST IPI\r\n"
    b"   GridNotasItens.ColWidth(30) = IIf(bIPI, 850, 0)  '%IPI\r\n"
    b"   GridNotasItens.ColWidth(31) = IIf(bIPI, 850, 0)  'IPI\r\n"
    b"End Sub\r\n",

    b"Sub AplicarVisibilidadeGridItens()\r\n"
    b"   If GridNotasItens.Cols < 34 Then Exit Sub\r\n"
    b"   'Reforma tributaria CBS/IBS: chkReforma\r\n"
    b"   Dim bReforma As Boolean\r\n"
    b"   bReforma = (chkReforma.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(9) = IIf(bReforma, 700, 0)     'CST IBS/CBS\r\n"
    b"   GridNotasItens.ColWidth(10) = IIf(bReforma, 1200, 0)   'cClassTrib\r\n"
    b"   GridNotasItens.ColWidth(13) = IIf(bReforma, 850, 0)    'V. IBS\r\n"
    b"   GridNotasItens.ColWidth(14) = IIf(bReforma, 850, 0)    'V. CBS\r\n"
    b"   GridNotasItens.ColWidth(15) = IIf(bReforma, 850, 0)    'V. IS\r\n"
    b"\r\n"
    b"   'Reforma tributaria IS: chkReformaIS\r\n"
    b"   Dim bReformaIS As Boolean\r\n"
    b"   bReformaIS = (chkReformaIS.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(11) = IIf(bReformaIS, 700, 0)   'CST IS\r\n"
    b"   GridNotasItens.ColWidth(12) = IIf(bReformaIS, 1200, 0)  'CLASS IS\r\n"
    b"\r\n"
    b"   'Seguro: chkSeguro\r\n"
    b"   GridNotasItens.ColWidth(19) = IIf(chkSeguro.Value = 1, 900, 0)\r\n"
    b"   'Outros: chkOutros\r\n"
    b"   GridNotasItens.ColWidth(20) = IIf(chkOutros.Value = 1, 900, 0)\r\n"
    b"\r\n"
    b"   'Grupo ICMS: exibe quando finalidade = 4 (devolucao/retorno)\r\n"
    b"   Dim bICMS As Boolean\r\n"
    b"   bICMS = (Left(cboFinalidade.Text, 1) = \"4\")\r\n"
    b"   GridNotasItens.ColWidth(23) = IIf(bICMS, 850, 0)  'BC ICMS\r\n"
    b"   GridNotasItens.ColWidth(24) = IIf(bICMS, 850, 0)  '%ICMS\r\n"
    b"   GridNotasItens.ColWidth(25) = IIf(bICMS, 850, 0)  'ICMS\r\n"
    b"\r\n"
    b"   '%RedBC: chkpRedBC\r\n"
    b"   GridNotasItens.ColWidth(26) = IIf(chkpRedBC.Value = 1, 700, 0)\r\n"
    b"\r\n"
    b"   'Grupo ICMSST: chkICMSST\r\n"
    b"   Dim bST As Boolean\r\n"
    b"   bST = (chkICMSST.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(27) = IIf(bST, 850, 0)  'BC ST\r\n"
    b"   GridNotasItens.ColWidth(28) = IIf(bST, 900, 0)  '%ICMSST\r\n"
    b"   GridNotasItens.ColWidth(29) = IIf(bST, 850, 0)  'ICMSST\r\n"
    b"   GridNotasItens.ColWidth(30) = IIf(bST, 850, 0)  'MVA ST\r\n"
    b"\r\n"
    b"   'Grupo IPI: chkIPI\r\n"
    b"   Dim bIPI As Boolean\r\n"
    b"   bIPI = (chkIPI.Value = 1)\r\n"
    b"   GridNotasItens.ColWidth(31) = IIf(bIPI, 850, 0)  'CST IPI\r\n"
    b"   GridNotasItens.ColWidth(32) = IIf(bIPI, 850, 0)  '%IPI\r\n"
    b"   GridNotasItens.ColWidth(33) = IIf(bIPI, 850, 0)  'IPI\r\n"
    b"End Sub\r\n"
))

# ── P10: Adicionar chkReformaIS_Click após chkReforma_Click ──────────────────
patches.append((
    b"Sub chkReforma_Click()\r\n"
    b"    AplicarVisibilidadeGridItens\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Sub chkSeguro_Click()\r\n",

    b"Sub chkReforma_Click()\r\n"
    b"    AplicarVisibilidadeGridItens\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Sub chkReformaIS_Click()\r\n"
    b"    AplicarVisibilidadeGridItens\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Sub chkSeguro_Click()\r\n"
))

# ── P11: Exibir_Itens — adicionar IS_CST e cClassTrib_IS ao SELECT ────────────
patches.append((
    b"       \"IBSCBS_CST, cClassTrib, IBS_vIBS, CBS_vCBS, IS_vIS, \" & _\r\n",
    b"       \"IBSCBS_CST, cClassTrib, IS_CST, cClassTrib_IS, IBS_vIBS, CBS_vCBS, IS_vIS, \" & _\r\n"
))

# ── P12: GridNotasItens_Click — Case list ─────────────────────────────────────
patches.append((
    b"    Case 2, 5, 6, 7, 8, 9, 10, 22, 24, 26, 28, 29, 30",
    b"    Case 2, 5, 6, 7, 8, 9, 10, 24, 26, 28, 30, 31, 32"
))

# ── P13: Case 9 grid updates — col 11→13, 12→14 ──────────────────────────────
patches.append((
    b"        GridNotasItens.TextMatrix(iRow, 9) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 10) = sNewClassTrib\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 11) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 12) = FormatNumber(curCBSvCBS, 2)\r\n",

    b"        GridNotasItens.TextMatrix(iRow, 9) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 10) = sNewClassTrib\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
))

# ── P14: Case 10 grid updates — col 11→13, 12→14 ─────────────────────────────
patches.append((
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 11) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 12) = FormatNumber(curCBSvCBS, 2)\r\n",

    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
))

# ── P15: Case 22 → Case 24 (col refs 21→23, 22→24, 23→25) ───────────────────
patches.append((
    b"    Case 22 ' %ICMS\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota ICMS inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPICMS = Val(sVal)\r\n"
    b"        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 21), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMS = CCur(Format(curVBC * dblPICMS / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pICMS = \" & FSQL(dblPICMS, 4) & \", vICMS = \" & FSQL(curVICMS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 22) = FormatNumber(dblPICMS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(curVICMS, 2)\r\n",

    b"    Case 24 ' %ICMS\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota ICMS inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPICMS = Val(sVal)\r\n"
    b"        curVBC = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 23), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMS = CCur(Format(curVBC * dblPICMS / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pICMS = \" & FSQL(dblPICMS, 4) & \", vICMS = \" & FSQL(curVICMS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 24) = FormatNumber(dblPICMS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(curVICMS, 2)\r\n"
))

# ── P16: Case 24 → Case 26 (col refs 20→22, 22→24, 21→23, 23→25, 24→26) ─────
patches.append((
    b"    Case 24 ' %RED BC\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPRedBC = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), \"0.00\"))\r\n"
    b"        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), \".\", \"\"), \",\", \".\"))\r\n"
    b"        curVICMS = CCur(Format(curVBC * dblPICMS / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pRedBC = \" & FSQL(dblPRedBC, 4) & \", vBC = \" & FSQL(curVBC, 2) & \", vICMS = \" & FSQL(curVICMS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 24) = FormatNumber(dblPRedBC, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 21) = FormatNumber(curVBC, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(curVICMS, 2)\r\n",

    b"    Case 26 ' %RED BC\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Redu\xe7\xe3o BC inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPRedBC = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVBC = CCur(Format(curSubTot * (1 - dblPRedBC / 100), \"0.00\"))\r\n"
    b"        dblPICMS = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 24), \".\", \"\"), \",\", \".\"))\r\n"
    b"        curVICMS = CCur(Format(curVBC * dblPICMS / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pRedBC = \" & FSQL(dblPRedBC, 4) & \", vBC = \" & FSQL(curVBC, 2) & \", vICMS = \" & FSQL(curVICMS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(dblPRedBC, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 23) = FormatNumber(curVBC, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(curVICMS, 2)\r\n"
))

# ── P17: Case 26 → Case 28 (col refs 25→27, 23→25, 26→28, 27→29) ────────────
patches.append((
    b"    Case 26 ' %ICMSST\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPICMSST = Val(sVal)\r\n"
    b"        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 23), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, \"0.00\")) - curVICMS\r\n"
    b"        If curVICMSST < 0 Then curVICMSST = 0\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pICMSST = \" & FSQL(dblPICMSST, 4) & \", vICMSST = \" & FSQL(curVICMSST, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 26) = FormatNumber(dblPICMSST, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(curVICMSST, 2)\r\n",

    b"    Case 28 ' %ICMSST\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota ICMS-ST inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPICMSST = Val(sVal)\r\n"
    b"        curVBCST = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 27), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, \"0.00\")) - curVICMS\r\n"
    b"        If curVICMSST < 0 Then curVICMSST = 0\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pICMSST = \" & FSQL(dblPICMSST, 4) & \", vICMSST = \" & FSQL(curVICMSST, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 28) = FormatNumber(dblPICMSST, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 29) = FormatNumber(curVICMSST, 2)\r\n"
))

# ── P18: Case 28 → Case 30 (col refs 20→22, 31→33, 26→28, 23→25, 28→30, 25→27, 27→29)
patches.append((
    b"    Case 28 ' MVA ST\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then\r\n"
    b"            MsgBox \"MVA inv\xe1lido (deve ser >= 0)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblMVA = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 31), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), \"0.00\"))\r\n"
    b"        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 26), \".\", \"\"), \",\", \".\"))\r\n"
    b"        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 23), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, \"0.00\")) - curVICMS\r\n"
    b"        If curVICMSST < 0 Then curVICMSST = 0\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pMVAST = \" & FSQL(dblMVA, 4) & \", vBCST = \" & FSQL(curVBCST, 2) & \", vICMSST = \" & FSQL(curVICMSST, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 28) = FormatNumber(dblMVA, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 25) = FormatNumber(curVBCST, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(curVICMSST, 2)\r\n",

    b"    Case 30 ' MVA ST\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Then\r\n"
    b"            MsgBox \"MVA inv\xe1lido (deve ser >= 0)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblMVA = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVIPI = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 33), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVBCST = CCur(Format((curSubTot + curVIPI) * (1 + dblMVA / 100), \"0.00\"))\r\n"
    b"        dblPICMSST = Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 28), \".\", \"\"), \",\", \".\"))\r\n"
    b"        curVICMS = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 25), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVICMSST = CCur(Format(curVBCST * dblPICMSST / 100, \"0.00\")) - curVICMS\r\n"
    b"        If curVICMSST < 0 Then curVICMSST = 0\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET pMVAST = \" & FSQL(dblMVA, 4) & \", vBCST = \" & FSQL(curVBCST, 2) & \", vICMSST = \" & FSQL(curVICMSST, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 30) = FormatNumber(dblMVA, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 27) = FormatNumber(curVBCST, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 29) = FormatNumber(curVICMSST, 2)\r\n"
))

# ── P19: Case 29 → Case 31 ────────────────────────────────────────────────────
patches.append((
    b"    Case 29 ' CST IPI\r\n"
    b"        If sVal = \"\" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then\r\n"
    b"            MsgBox \"CST IPI deve ter 2 d\xedgitos!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IPICST = '\" & sVal & \"', IPIcEnq = '999' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n",

    b"    Case 31 ' CST IPI\r\n"
    b"        If sVal = \"\" Or Len(sVal) <> 2 Or Not IsNumeric(sVal) Then\r\n"
    b"            MsgBox \"CST IPI deve ter 2 d\xedgitos!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IPICST = '\" & sVal & \"', IPIcEnq = '999' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
))

# ── P20: Case 30 → Case 32 (col refs 20→22, 30→32, 31→33) ───────────────────
patches.append((
    b"    Case 30 ' %IPI\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota IPI inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPIPI = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 20), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IPIpIPI = \" & FSQL(dblPIPI, 4) & \", IPIvIPI = \" & FSQL(curVIPI, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 30) = FormatNumber(dblPIPI, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 31) = FormatNumber(curVIPI, 2)\r\n",

    b"    Case 32 ' %IPI\r\n"
    b"        sVal = Replace(Replace(sVal, \".\", \"\"), \",\", \".\")\r\n"
    b"        If Not IsNumeric(sVal) Or Val(sVal) < 0 Or Val(sVal) > 100 Then\r\n"
    b"            MsgBox \"Al\xedquota IPI inv\xe1lida (0 a 100)!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPIPI = Val(sVal)\r\n"
    b"        curSubTot = CCur(Val(Replace(Replace(GridNotasItens.TextMatrix(iRow, 22), \".\", \"\"), \",\", \".\")))\r\n"
    b"        curVIPI = CCur(Format(curSubTot * dblPIPI / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IPIpIPI = \" & FSQL(dblPIPI, 4) & \", IPIvIPI = \" & FSQL(curVIPI, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 32) = FormatNumber(dblPIPI, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 33) = FormatNumber(curVIPI, 2)\r\n"
))

# ── Aplicar ───────────────────────────────────────────────────────────────────
errors = 0
for idx, (old, new) in enumerate(patches, 1):
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO P{idx}: count={cnt} (esperado 1)")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   P{idx}")

data = norm(data)

if errors:
    print(f"\n{errors} erro(s). Arquivo NÃO foi salvo.")
    sys.exit(1)

with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
