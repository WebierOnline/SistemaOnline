import re

# ============================================================
# 1. Modificar cmdCadastrar_Click em frmVinculoProdutoXML.frm
# ============================================================
data = open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'rb').read()

old_sub = (
    b"Sub cmdCadastrar_Click()\r\n"
    b"   If iSelecionado < 0 Then\r\n"
    b"      MsgBox \"Selecione um item da lista.\", vbExclamation\r\n"
    b"      Exit Sub\r\n"
    b"   End If\r\n"
    b"\r\n"
    b"   Dim frac As Double\r\n"
    b"   frac = Val(txtFracionamento.Text)\r\n"
    b"   If frac <= 0 Then frac = 1\r\n"
    b"\r\n"
    b"   Dim item As tItemXML\r\n"
    b"   item = arrItens(iSelecionado)\r\n"
    b"\r\n"
    b"   'Pergunta modo de venda\r\n"
    b"   Dim resp As Integer\r\n"
    b"   resp = MsgBox(\"Como este produto sera vendido?\" & vbCrLf & vbCrLf & _\r\n"
    b"                 \"SIM  = Atacado (vende a caixa/embalagem do fornecedor)\" & vbCrLf & _\r\n"
    b"                 \"NAO  = Varejo  (vende a unidade individual)\", _\r\n"
    b"                 vbQuestion + vbYesNoCancel, \"Modo de Venda\")\r\n"
    b"   If resp = vbCancel Then Exit Sub\r\n"
    b"   Dim bVarejo As Boolean\r\n"
    b"   bVarejo = (resp = vbNo)\r\n"
)

new_sub_start = (
    b"Sub cmdCadastrar_Click()\r\n"
    b"   If iSelecionado < 0 Then\r\n"
    b"      MsgBox \"Selecione um item da lista.\", vbExclamation\r\n"
    b"      Exit Sub\r\n"
    b"   End If\r\n"
    b"\r\n"
    b"   Dim frac As Double\r\n"
    b"   frac = Val(txtFracionamento.Text)\r\n"
    b"   If frac <= 0 Then frac = 1\r\n"
    b"\r\n"
    b"   Dim item As tItemXML\r\n"
    b"   item = arrItens(iSelecionado)\r\n"
    b"\r\n"
    b"   ' Exibir form de escolha do modo de cadastro\r\n"
    b"   Load frmModoVenda\r\n"
    b"   frmModoVenda.SetNome item.Nome\r\n"
    b"   frmModoVenda.Show vbModal, Me\r\n"
    b"   Dim nEscolha As Integer\r\n"
    b"   nEscolha = frmModoVenda.Escolha\r\n"
    b"   Unload frmModoVenda\r\n"
    b"   If nEscolha = 0 Then Exit Sub\r\n"
    b"\r\n"
    b"   Dim bVarejo As Boolean\r\n"
    b"   bVarejo = (nEscolha = 2)\r\n"
)

c = data.count(old_sub)
print(f'1. Cabecalho cmdCadastrar_Click: count={c}')
if c == 1:
    data = data.replace(old_sub, new_sub_start)

# 2. Adicionar bloco MANUAL antes do bloco de finalizacao
# O bloco ATACADO termina com: sEANCad = item.sEAN\r\n   End If\r\n\r\n   'Sanitiza
old_atacado_end = (
    b"      sDesc = item.Nome\r\n"
    b"      sEANCad = item.sEAN\r\n"
    b"   End If\r\n"
    b"\r\n"
    b"   'Sanitiza\r\n"
)

new_atacado_end = (
    b"      sDesc = item.Nome\r\n"
    b"      sEANCad = item.sEAN\r\n"
    b"   End If\r\n"
    b"\r\n"
    b"   ' MANUAL: abre Produtos_Cadastro pre-preenchido com dados da XML\r\n"
    b"   If nEscolha = 3 Then\r\n"
    b"      Dim lMaxCodAnt As Long\r\n"
    b"      lMaxCodAnt = SQLExecutaRetorno(\"SELECT ISNULL(MAX(CODIGO),0) r FROM Produtos\", \"r\", 0)\r\n"
    b"\r\n"
    b"      Load Produtos_Cadastro\r\n"
    b"      Produtos_Cadastro.SSTab1.Tab = 0\r\n"
    b"      Produtos_Cadastro.CriarNovoProduto\r\n"
    b"\r\n"
    b"      ' Dados basicos\r\n"
    b"      If item.sEAN = \"SEM GTIN\" Or item.sEAN = \"\" Then\r\n"
    b"          Produtos_Cadastro.txtCodBarra.Text = \"\"\r\n"
    b"          Produtos_Cadastro.txtEAN.Text = \"\"\r\n"
    b"      Else\r\n"
    b"          Produtos_Cadastro.txtCodBarra.Text = item.sEAN\r\n"
    b"          Produtos_Cadastro.txtEAN.Text = item.sEAN\r\n"
    b"      End If\r\n"
    b"      Produtos_Cadastro.txtDescricao.Text = item.Nome\r\n"
    b"      Produtos_Cadastro.cboUnidMedida.Text = ConverterUnidade(item.uCom, False)\r\n"
    b"      Produtos_Cadastro.txtNCM.Text = item.NCM\r\n"
    b"      Produtos_Cadastro.txtCEST.Text = item.CEST\r\n"
    b"\r\n"
    b"      ' CFOP e CST ja convertidos pelo regime\r\n"
    b"      Produtos_Cadastro.cboCFOP.Text = sCFOPSaida\r\n"
    b"      Produtos_Cadastro.cboCST.Text = sICMSCST\r\n"
    b"\r\n"
    b"      ' ICMS\r\n"
    b"      Produtos_Cadastro.txtICMSAliquota.Text = FormatNumber(dICMSAliq, 2)\r\n"
    b"      Produtos_Cadastro.txtRedBCAliquota.Text = FormatNumber(dpRedBC, 2)\r\n"
    b"\r\n"
    b"      ' ST (nao disponivel em tItemXML - usuario completa)\r\n"
    b"      Produtos_Cadastro.txtMVA.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.txtSTAliq.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.txtRedBCST.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.cboModBC.ListIndex = 3\r\n"
    b"      Produtos_Cadastro.cboModBCST.ListIndex = 4\r\n"
    b"\r\n"
    b"      ' PIS / COFINS / IPI\r\n"
    b"      Produtos_Cadastro.txtPISCST.Text = sPISCST\r\n"
    b"      Produtos_Cadastro.txtPisAliquota.Text = FormatNumber(dPISAliq, 2)\r\n"
    b"      Produtos_Cadastro.txtCOFINSCST.Text = sCOFINSCST\r\n"
    b"      Produtos_Cadastro.txtCofinsAliquota.Text = FormatNumber(dCOFINSAliq, 2)\r\n"
    b"      Produtos_Cadastro.txtIPICST.Text = sIPICST\r\n"
    b"      Produtos_Cadastro.txtIPIAliquota.Text = FormatNumber(dIPIAliq, 2)\r\n"
    b"\r\n"
    b"      ' Reforma Tributaria: selecionar combos pelo prefixo de 2 digitos\r\n"
    b"      Dim kM As Integer\r\n"
    b"      For kM = 0 To Produtos_Cadastro.cboIBSCBSCST.ListCount - 1\r\n"
    b"          If Left(Produtos_Cadastro.cboIBSCBSCST.List(kM), 2) = \"01\" Then\r\n"
    b"              Produtos_Cadastro.cboIBSCBSCST.ListIndex = kM: Exit For\r\n"
    b"          End If\r\n"
    b"      Next kM\r\n"
    b"      Produtos_Cadastro.txtCBSpAliq.Text = \"0,0000\"\r\n"
    b"      Produtos_Cadastro.txtIBSUFpAliq.Text = \"0,0000\"\r\n"
    b"      Produtos_Cadastro.txtIBSMunpAliq.Text = \"0,0000\"\r\n"
    b"      For kM = 0 To Produtos_Cadastro.cboISCST.ListCount - 1\r\n"
    b"          If Left(Produtos_Cadastro.cboISCST.List(kM), 2) = \"00\" Then\r\n"
    b"              Produtos_Cadastro.cboISCST.ListIndex = kM: Exit For\r\n"
    b"          End If\r\n"
    b"      Next kM\r\n"
    b"      Produtos_Cadastro.txtISpIS.Text = \"0,0000\"\r\n"
    b"\r\n"
    b"      ' Custo e margens\r\n"
    b"      Produtos_Cadastro.txtCusto.Text = Format(item.vUnCom / frac, \"##,##0.00\")\r\n"
    b"      Produtos_Cadastro.txtMargemVV.Text = \"0,00%\"\r\n"
    b"      Produtos_Cadastro.txtMargemVP.Text = \"0,00%\"\r\n"
    b"      Produtos_Cadastro.txtMargemAV.Text = \"0,00%\"\r\n"
    b"      Produtos_Cadastro.txtMargemAP.Text = \"0,00%\"\r\n"
    b"      Produtos_Cadastro.txtValorVV.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.txtValorVP.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.txtValorAV.Text = \"0,00\"\r\n"
    b"      Produtos_Cadastro.txtValorAP.Text = \"0,00\"\r\n"
    b"\r\n"
    b"      Produtos_Cadastro.Show vbModal, Me\r\n"
    b"\r\n"
    b"      ' Verificar se produto foi salvo (codigo maior que o anterior)\r\n"
    b"      Dim lNovoCod As Long\r\n"
    b"      lNovoCod = SQLExecutaRetorno(\"SELECT ISNULL(MAX(CODIGO),0) r FROM Produtos WHERE CODIGO > \" & lMaxCodAnt, \"r\", 0)\r\n"
    b"      If lNovoCod > 0 Then\r\n"
    b"          If Not ExecutarVinculo(lNovoCod, frac) Then Exit Sub\r\n"
    b"          arrItens(iSelecionado).Vinculado = True\r\n"
    b"          arrItens(iSelecionado).IDProdVinculado = lNovoCod\r\n"
    b"          AtualizarItemLista iSelecionado\r\n"
    b"          AvancarParaProximo\r\n"
    b"          VerificarConclusao\r\n"
    b"      Else\r\n"
    b"          MsgBox \"Produto nao salvo. Realize o vinculo manualmente apos o cadastro.\", vbInformation\r\n"
    b"      End If\r\n"
    b"      Exit Sub\r\n"
    b"   End If\r\n"
    b"\r\n"
    b"   'Sanitiza\r\n"
)

c = data.count(old_atacado_end)
print(f'2. Bloco MANUAL antes de Sanitiza: count={c}')
if c == 1:
    data = data.replace(old_atacado_end, new_atacado_end)

data = data.replace(b'\r\n', b'\n').replace(b'\r', b'\n').replace(b'\n', b'\r\n')
open('OnlineCommerce/Forms/frmVinculoProdutoXML.frm', 'wb').write(data)
print('frmVinculoProdutoXML salvo. Tamanho:', len(data))

# ============================================================
# 3. Adicionar frmModoVenda ao OnlineCommerce.vbp
# ============================================================
vbp = open('OnlineCommerce/OnlineCommerce.vbp', 'rb').read()
old_ref = b'Form=Forms\\frmCadProdXML.frm\r\n'
new_ref = b'Form=Forms\\frmCadProdXML.frm\r\nForm=Forms\\frmModoVenda.frm\r\n'
c = vbp.count(old_ref)
print(f'3. Adicionar frmModoVenda ao .vbp: count={c}')
if c == 1:
    vbp = vbp.replace(old_ref, new_ref)
    open('OnlineCommerce/OnlineCommerce.vbp', 'wb').write(vbp)
    print('.vbp salvo.')
