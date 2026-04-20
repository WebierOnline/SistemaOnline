data = open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'rb').read()

# =============================================================================
# 1. tItemXML: adicionar campos de ST, modBC/ST, IBS/CBS, IS
# =============================================================================
old_type = (
    b"Private Type tItemXML\r\n"
    b"   cProd      As String\r\n"
    b"   sEAN       As String\r\n"
    b"   Nome       As String\r\n"
    b"   uCom       As String\r\n"
    b"   vUnCom     As Double\r\n"
    b"   NCM        As String\r\n"
    b"   CEST       As String\r\n"
    b"   Vinculado       As Boolean\r\n"
    b"   IDProdVinculado As Long      '0 = nao vinculado\r\n"
    b"   '--- Tributacao da NF-e (entrada) ---\r\n"
    b"   ICMSCST    As String\r\n"
    b"   ICMSAliq   As Double\r\n"
    b"   pRedBC     As Double\r\n"
    b"   IPICST     As String\r\n"
    b"   IPIAliq    As Double\r\n"
    b"   PISCST     As String\r\n"
    b"   PISAliq    As Double\r\n"
    b"   COFINSCST  As String\r\n"
    b"   COFINSAliq As Double\r\n"
    b"   CFOP       As String\r\n"
    b"End Type\r\n"
)
new_type = (
    b"Private Type tItemXML\r\n"
    b"   cProd      As String\r\n"
    b"   sEAN       As String\r\n"
    b"   Nome       As String\r\n"
    b"   uCom       As String\r\n"
    b"   vUnCom     As Double\r\n"
    b"   NCM        As String\r\n"
    b"   CEST       As String\r\n"
    b"   Vinculado       As Boolean\r\n"
    b"   IDProdVinculado As Long\r\n"
    b"   '--- Tributacao ICMS ---\r\n"
    b"   ICMSCST    As String\r\n"
    b"   ICMSAliq   As Double\r\n"
    b"   pRedBC     As Double\r\n"
    b"   modBC      As Integer\r\n"
    b"   '--- Substituicao Tributaria ---\r\n"
    b"   pMVAST     As Double\r\n"
    b"   pICMSST    As Double\r\n"
    b"   pRedBCST   As Double\r\n"
    b"   modBCST    As Integer\r\n"
    b"   '--- IPI / PIS / COFINS ---\r\n"
    b"   IPICST     As String\r\n"
    b"   IPIAliq    As Double\r\n"
    b"   PISCST     As String\r\n"
    b"   PISAliq    As Double\r\n"
    b"   COFINSCST  As String\r\n"
    b"   COFINSAliq As Double\r\n"
    b"   CFOP       As String\r\n"
    b"   '--- Reforma Tributaria (IBS/CBS/IS) ---\r\n"
    b"   IBSCBSCST  As String\r\n"
    b"   IBSUFpAliq As Double\r\n"
    b"   IBSMunpAliq As Double\r\n"
    b"   CBSpAliq   As Double\r\n"
    b"   ISCST      As String\r\n"
    b"   ISpIS      As Double\r\n"
    b"End Type\r\n"
)
c = data.count(old_type)
print(f'1. tItemXML: count={c}')
if c == 1: data = data.replace(old_type, new_type)

# =============================================================================
# 2. Form_Load: SELECT — adicionar campos novos
# =============================================================================
old_select = (
    b'   RsOpen rs, "SELECT Item, Referencia AS cProd, ISNULL(EAN,\'\') AS EAN, " & _\r\n'
    b'              "       NomeProduto, UnidadeComercial AS uCom, " & _\r\n'
    b'              "       ValorUnitarioComercializacao AS vUnCom, " & _\r\n'
    b'              "       ISNULL(NCM,\'\') AS NCM, ISNULL(CEST,\'\') AS CEST, " & _\r\n'
    b'              "       ISNULL(CST,\'\') AS ICMSCST, ISNULL(pICMS,0) AS ICMSAliq, " & _\r\n'
    b'              "       ISNULL(pRedBC,0) AS pRedBC, " & _\r\n'
    b'              "       ISNULL(IPICST,\'\') AS IPICST, ISNULL(IPIpIPI,0) AS IPIAliq, " & _\r\n'
    b'              "       ISNULL(pisCST,\'\') AS PISCST, ISNULL(PISpPIS,0) AS PISAliq, " & _\r\n'
    b'              "       ISNULL(cofinsCST,\'\') AS COFINSCST, " & _\r\n'
    b'              "       ISNULL(COFINSpCOFINS,0) AS COFINSAliq, " & _\r\n'
    b'              "       ISNULL(CFOP,\'\') AS CFOP, " & _\r\n'
    b'              "       ISNULL(CodigoProduto, 0) AS IDProdVinculado, " & _\r\n'
    b'              "       CASE WHEN (CodigoProduto IS NULL OR CodigoProduto = 0) " & _\r\n'
    b'              "            THEN 0 ELSE 1 END AS jaVinculado " & _\r\n'
    b'              "FROM EntradaEstoqueItens " & sWhere & " ORDER BY Item"\r\n'
)
new_select = (
    b'   RsOpen rs, "SELECT Item, Referencia AS cProd, ISNULL(EAN,\'\') AS EAN, " & _\r\n'
    b'              "       NomeProduto, UnidadeComercial AS uCom, " & _\r\n'
    b'              "       ValorUnitarioComercializacao AS vUnCom, " & _\r\n'
    b'              "       ISNULL(NCM,\'\') AS NCM, ISNULL(CEST,\'\') AS CEST, " & _\r\n'
    b'              "       ISNULL(CST,\'\') AS ICMSCST, ISNULL(pICMS,0) AS ICMSAliq, " & _\r\n'
    b'              "       ISNULL(pRedBC,0) AS pRedBC, ISNULL(modBC,3) AS modBC, " & _\r\n'
    b'              "       ISNULL(pMVAST,0) AS pMVAST, ISNULL(pICMSST,0) AS pICMSST, " & _\r\n'
    b'              "       ISNULL(pRedBCST,0) AS pRedBCST, ISNULL(modBCST,4) AS modBCST, " & _\r\n'
    b'              "       ISNULL(IPICST,\'\') AS IPICST, ISNULL(IPIpIPI,0) AS IPIAliq, " & _\r\n'
    b'              "       ISNULL(pisCST,\'\') AS PISCST, ISNULL(PISpPIS,0) AS PISAliq, " & _\r\n'
    b'              "       ISNULL(cofinsCST,\'\') AS COFINSCST, " & _\r\n'
    b'              "       ISNULL(COFINSpCOFINS,0) AS COFINSAliq, " & _\r\n'
    b'              "       ISNULL(CFOP,\'\') AS CFOP, " & _\r\n'
    b'              "       ISNULL(IBSCBSCST,\'\') AS IBSCBSCST, " & _\r\n'
    b'              "       ISNULL(IBSUFpAliq,0) AS IBSUFpAliq, ISNULL(IBSMunpAliq,0) AS IBSMunpAliq, " & _\r\n'
    b'              "       ISNULL(CBSpAliq,0) AS CBSpAliq, " & _\r\n'
    b'              "       ISNULL(ISCST,\'\') AS ISCST, ISNULL(ISpIS,0) AS ISpIS, " & _\r\n'
    b'              "       ISNULL(CodigoProduto, 0) AS IDProdVinculado, " & _\r\n'
    b'              "       CASE WHEN (CodigoProduto IS NULL OR CodigoProduto = 0) " & _\r\n'
    b'              "            THEN 0 ELSE 1 END AS jaVinculado " & _\r\n'
    b'              "FROM EntradaEstoqueItens " & sWhere & " ORDER BY Item"\r\n'
)
c = data.count(old_select)
print(f'2. Form_Load SELECT: count={c}')
if c == 1: data = data.replace(old_select, new_select)

# =============================================================================
# 3. Form_Load: loop de populacao — adicionar leitura dos campos novos
# =============================================================================
old_loop = (
    b'         .vUnCom = 0\r\n'
    b'         .ICMSAliq = 0\r\n'
    b'         .pRedBC = 0\r\n'
    b'         .IPIAliq = 0\r\n'
    b'         .PISAliq = 0\r\n'
    b'         .COFINSAliq = 0\r\n'
    b'         .Vinculado = (rs!jaVinculado <> 0)\r\n'
    b'         .IDProdVinculado = 0\r\n'
    b'         On Error Resume Next\r\n'
    b'         .IDProdVinculado = CLng(rs!IDProdVinculado)\r\n'
    b'         On Error GoTo ErrForm_Load\r\n'
    b'         On Error Resume Next\r\n'
    b'         .vUnCom = CDbl(rs!vUnCom)\r\n'
    b'         .ICMSAliq = CDbl(rs!ICMSAliq)\r\n'
    b'         .pRedBC = CDbl(rs!pRedBC)\r\n'
    b'         .IPIAliq = CDbl(rs!IPIAliq)\r\n'
    b'         .PISAliq = CDbl(rs!PISAliq)\r\n'
    b'         .COFINSAliq = CDbl(rs!COFINSAliq)\r\n'
    b'         On Error GoTo ErrForm_Load\r\n'
)
new_loop = (
    b'         .vUnCom = 0\r\n'
    b'         .ICMSAliq = 0: .pRedBC = 0: .modBC = 3\r\n'
    b'         .pMVAST = 0: .pICMSST = 0: .pRedBCST = 0: .modBCST = 4\r\n'
    b'         .IPIAliq = 0: .PISAliq = 0: .COFINSAliq = 0\r\n'
    b'         .IBSUFpAliq = 0: .IBSMunpAliq = 0: .CBSpAliq = 0: .ISpIS = 0\r\n'
    b'         .IBSCBSCST = rs!IBSCBSCST & "": .ISCST = rs!ISCST & ""\r\n'
    b'         .Vinculado = (rs!jaVinculado <> 0)\r\n'
    b'         .IDProdVinculado = 0\r\n'
    b'         On Error Resume Next\r\n'
    b'         .IDProdVinculado = CLng(rs!IDProdVinculado)\r\n'
    b'         On Error GoTo ErrForm_Load\r\n'
    b'         On Error Resume Next\r\n'
    b'         .vUnCom = CDbl(rs!vUnCom)\r\n'
    b'         .ICMSAliq = CDbl(rs!ICMSAliq)\r\n'
    b'         .pRedBC = CDbl(rs!pRedBC): .modBC = CInt(rs!modBC)\r\n'
    b'         .pMVAST = CDbl(rs!pMVAST): .pICMSST = CDbl(rs!pICMSST)\r\n'
    b'         .pRedBCST = CDbl(rs!pRedBCST): .modBCST = CInt(rs!modBCST)\r\n'
    b'         .IPIAliq = CDbl(rs!IPIAliq)\r\n'
    b'         .PISAliq = CDbl(rs!PISAliq): .COFINSAliq = CDbl(rs!COFINSAliq)\r\n'
    b'         .IBSUFpAliq = CDbl(rs!IBSUFpAliq): .IBSMunpAliq = CDbl(rs!IBSMunpAliq)\r\n'
    b'         .CBSpAliq = CDbl(rs!CBSpAliq): .ISpIS = CDbl(rs!ISpIS)\r\n'
    b'         On Error GoTo ErrForm_Load\r\n'
)
c = data.count(old_loop)
print(f'3. Form_Load loop: count={c}')
if c == 1: data = data.replace(old_loop, new_loop)

# =============================================================================
# 4. INSERT ATACADO/VAREJO: adicionar campos novos
# =============================================================================
old_insert = (
    b'   sSQL = "INSERT INTO Produtos " & _\r\n'
    b'          "(codigo, ativo, destaque, USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO, " & _\r\n'
    b'          " cod_barra, ean, descricao, fabricante, unid_medida, categoria, PRATELEIRA, " & _\r\n'
    b'          " quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, " & _\r\n'
    b'          " ICMSCST, ICMSAliq, PISCST, PISALIQ, COFINSCST, COFINSALIQ, IPICST, IPIALIQ, pRedBc, " & _\r\n'
    b'          " NCM, CEST, CFOP, Alterado, PedirPeso, CODPROD_FRACAO, QUANT_FRACAO) " & _\r\n'
    b'          "VALUES (" & _\r\n'
    b'          novoCodigo & ", 1, 0, 0, 0, 0, 0, 0, " & _\r\n'
    b'          "\'\" & sEANCad & "\', \'\" & sEANCad & "\', \'\" & sDesc & "\', \'\', \'\" & sUnidade & "\', \'\', \'\', " & _\r\n'
    b'          "0, \'\', 0, \'\', \'\', " & _\r\n'
    b'          "\'" & sICMSCST & "\', " & FSQL(dICMSAliq) & ", " & _\r\n'
    b'          "\'" & sPISCST & "\', " & FSQL(dPISAliq) & ", " & _\r\n'
    b'          "\'" & sCOFINSCST & "\', " & FSQL(dCOFINSAliq) & ", " & _\r\n'
    b'          "\'" & sIPICST & "\', " & FSQL(dIPIAliq) & ", " & _\r\n'
    b'          FSQL(dpRedBC) & ", " & _\r\n'
    b'          "\'" & item.NCM & "\', \'" & item.CEST & "\', \'" & sCFOPSaida & "\', 0, 0, 0, 0)"\r\n'
)
new_insert = (
    b'   sSQL = "INSERT INTO Produtos " & _\r\n'
    b'          "(codigo, ativo, destaque, USOCONSUMO, COMBUSTIVEL, MATERIAPRIMA, IMOBILIZADO, FRACIONADO, " & _\r\n'
    b'          " cod_barra, ean, descricao, fabricante, unid_medida, categoria, PRATELEIRA, " & _\r\n'
    b'          " quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, " & _\r\n'
    b'          " ICMSCST, ICMSAliq, PISCST, PISALIQ, COFINSCST, COFINSALIQ, IPICST, IPIALIQ, pRedBc, " & _\r\n'
    b'          " NCM, CEST, CFOP, Alterado, PedirPeso, CODPROD_FRACAO, QUANT_FRACAO, " & _\r\n'
    b'          " pMVAST, pICMSST, pRedBCST, modBC, modBCST, " & _\r\n'
    b'          " IBSCBSCST, CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISCST, ISpIS) " & _\r\n'
    b'          "VALUES (" & _\r\n'
    b'          novoCodigo & ", 1, 0, 0, 0, 0, 0, 0, " & _\r\n'
    b'          "\'\" & sEANCad & "\', \'\" & sEANCad & "\', \'\" & sDesc & "\', \'\', \'\" & sUnidade & "\', \'\', \'\', " & _\r\n'
    b'          "0, \'\', 0, \'\', \'\', " & _\r\n'
    b'          "\'" & sICMSCST & "\', " & FSQL(dICMSAliq) & ", " & _\r\n'
    b'          "\'" & sPISCST & "\', " & FSQL(dPISAliq) & ", " & _\r\n'
    b'          "\'" & sCOFINSCST & "\', " & FSQL(dCOFINSAliq) & ", " & _\r\n'
    b'          "\'" & sIPICST & "\', " & FSQL(dIPIAliq) & ", " & _\r\n'
    b'          FSQL(dpRedBC) & ", " & _\r\n'
    b'          "\'" & item.NCM & "\', \'" & item.CEST & "\', \'" & sCFOPSaida & "\', 0, 0, 0, 0, " & _\r\n'
    b'          FSQL(item.pMVAST) & ", " & FSQL(item.pICMSST) & ", " & FSQL(item.pRedBCST) & ", " & _\r\n'
    b'          item.modBC & ", " & item.modBCST & ", " & _\r\n'
    b'          "\'" & item.IBSCBSCST & "\', " & FSQL(item.CBSpAliq) & ", " & _\r\n'
    b'          FSQL(item.IBSUFpAliq) & ", " & FSQL(item.IBSMunpAliq) & ", " & _\r\n'
    b'          "\'" & item.ISCST & "\', " & FSQL(item.ISpIS) & ")"\r\n'
)
c = data.count(old_insert)
print(f'4. INSERT ATACADO/VAREJO: count={c}')
if c == 1: data = data.replace(old_insert, new_insert)

# =============================================================================
# 5. MANUAL: substituir defaults zerados pelos valores reais de item.*
# =============================================================================
old_manual_st = (
    b'      \' ST (nao disponivel em tItemXML - usuario completa)\r\n'
    b'      Produtos_Cadastro.txtMVA.Text = "0,00"\r\n'
    b'      Produtos_Cadastro.txtSTAliq.Text = "0,00"\r\n'
    b'      Produtos_Cadastro.txtRedBCST.Text = "0,00"\r\n'
    b'      Produtos_Cadastro.cboModBC.ListIndex = 3\r\n'
    b'      Produtos_Cadastro.cboModBCST.ListIndex = 4\r\n'
    b'\r\n'
    b'      \' PIS / COFINS / IPI\r\n'
)
new_manual_st = (
    b'      \' ST\r\n'
    b'      Produtos_Cadastro.txtMVA.Text = FormatNumber(item.pMVAST, 2)\r\n'
    b'      Produtos_Cadastro.txtSTAliq.Text = FormatNumber(item.pICMSST, 2)\r\n'
    b'      Produtos_Cadastro.txtRedBCST.Text = FormatNumber(item.pRedBCST, 2)\r\n'
    b'      If item.modBC >= 0 And item.modBC <= 3 Then Produtos_Cadastro.cboModBC.ListIndex = item.modBC\r\n'
    b'      If item.modBCST >= 0 And item.modBCST <= 6 Then Produtos_Cadastro.cboModBCST.ListIndex = item.modBCST\r\n'
    b'\r\n'
    b'      \' PIS / COFINS / IPI\r\n'
)
c = data.count(old_manual_st)
print(f'5. MANUAL ST: count={c}')
if c == 1: data = data.replace(old_manual_st, new_manual_st)

old_manual_ibs = (
    b'      \' Reforma Tributaria: selecionar combos pelo prefixo de 2 digitos\r\n'
    b'      Dim kM As Integer\r\n'
    b'      For kM = 0 To Produtos_Cadastro.cboIBSCBSCST.ListCount - 1\r\n'
    b'          If Left(Produtos_Cadastro.cboIBSCBSCST.List(kM), 2) = "01" Then\r\n'
    b'              Produtos_Cadastro.cboIBSCBSCST.ListIndex = kM: Exit For\r\n'
    b'          End If\r\n'
    b'      Next kM\r\n'
    b'      Produtos_Cadastro.txtCBSpAliq.Text = "0,0000"\r\n'
    b'      Produtos_Cadastro.txtIBSUFpAliq.Text = "0,0000"\r\n'
    b'      Produtos_Cadastro.txtIBSMunpAliq.Text = "0,0000"\r\n'
    b'      For kM = 0 To Produtos_Cadastro.cboISCST.ListCount - 1\r\n'
    b'          If Left(Produtos_Cadastro.cboISCST.List(kM), 2) = "00" Then\r\n'
    b'              Produtos_Cadastro.cboISCST.ListIndex = kM: Exit For\r\n'
    b'          End If\r\n'
    b'      Next kM\r\n'
    b'      Produtos_Cadastro.txtISpIS.Text = "0,0000"\r\n'
)
new_manual_ibs = (
    b'      \' Reforma Tributaria: selecionar combos pelo prefixo de 2 digitos\r\n'
    b'      Dim kM As Integer\r\n'
    b'      Dim sIBSM As String\r\n'
    b'      sIBSM = Left(item.IBSCBSCST & "  ", 2)\r\n'
    b'      If Trim(sIBSM) = "" Then sIBSM = "01"\r\n'
    b'      For kM = 0 To Produtos_Cadastro.cboIBSCBSCST.ListCount - 1\r\n'
    b'          If Left(Produtos_Cadastro.cboIBSCBSCST.List(kM), 2) = sIBSM Then\r\n'
    b'              Produtos_Cadastro.cboIBSCBSCST.ListIndex = kM: Exit For\r\n'
    b'          End If\r\n'
    b'      Next kM\r\n'
    b'      Produtos_Cadastro.txtCBSpAliq.Text = FormatNumber(item.CBSpAliq, 4)\r\n'
    b'      Produtos_Cadastro.txtIBSUFpAliq.Text = FormatNumber(item.IBSUFpAliq, 4)\r\n'
    b'      Produtos_Cadastro.txtIBSMunpAliq.Text = FormatNumber(item.IBSMunpAliq, 4)\r\n'
    b'      Dim sISM As String\r\n'
    b'      sISM = Left(item.ISCST & "  ", 2)\r\n'
    b'      If Trim(sISM) = "" Then sISM = "00"\r\n'
    b'      For kM = 0 To Produtos_Cadastro.cboISCST.ListCount - 1\r\n'
    b'          If Left(Produtos_Cadastro.cboISCST.List(kM), 2) = sISM Then\r\n'
    b'              Produtos_Cadastro.cboISCST.ListIndex = kM: Exit For\r\n'
    b'          End If\r\n'
    b'      Next kM\r\n'
    b'      Produtos_Cadastro.txtISpIS.Text = FormatNumber(item.ISpIS, 4)\r\n'
)
c = data.count(old_manual_ibs)
print(f'6. MANUAL IBS/CBS: count={c}')
if c == 1: data = data.replace(old_manual_ibs, new_manual_ibs)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'wb').write(data)
print('Salvo. Tamanho:', len(data))
