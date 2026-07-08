"""
Patch Produtos_Cadastro.frm:
- Remove globals gCBSpAliq/gIBSUFpAliq/gIBSMunpAliq
- Remove carga de aliquotas da Empresa/Cidade
- Adiciona cboTipoIS.AddItem (1/2/3)
- cboISClasse_Click: auto-preenche cboTipoIS
- Carga ao editar: cClassTrib_IS e tipo_calculo_is
- LimparObjetos: limpa cboISClasse e cboTipoIS
- INSERT: troca CBSpAliq/IBSUFpAliq/IBSMunpAliq/ISpIS por cClassTrib_IS/tipo_calculo_is
- UPDATE: idem
- PreencherReformaDoNCM: seleciona cboTipoIS
"""
FILE = r"C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm"
with open(FILE, "rb") as f:
    raw = f.read()
data = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")

errors = 0

def patch(old, new, label):
    global data, errors
    c = data.count(old)
    if c != 1:
        print(f"ERRO [{label}] ({c}x)")
        errors += 1
    else:
        data = data.replace(old, new, 1)

# 1. Remover globals
patch(
    b"Dim gCBSpAliq As Double\nDim gIBSUFpAliq As Double\nDim gIBSMunpAliq As Double\n",
    b"",
    "globals"
)

# 2. Remover bloco carga aliquotas Empresa/Cidade
patch(
    b"' Carrega aliquotas IBS/CBS da Empresa e da Cidade da empresa\n"
    b"Dim rEmp As ADODB.Recordset\n"
    b"RsOpen rEmp, \"SELECT E.CBSpAliq, C.IBSUFpAliq, C.IBSMunpAliq \" & _\n"
    b"             \"FROM Empresa E \" & _\n"
    b"             \"LEFT JOIN Cidade C ON CAST(C.CodigoMunicipio AS NVARCHAR(7)) = \" & _\n"
    b"             \"                     CAST(E.CodigoIBGE AS NVARCHAR(7))\"\n"
    b"If Not rEmp.EOF Then\n"
    b"    gCBSpAliq = IIf(IsNull(rEmp(\"CBSpAliq\")), 0, rEmp(\"CBSpAliq\"))\n"
    b"    gIBSUFpAliq = IIf(IsNull(rEmp(\"IBSUFpAliq\")), 0, rEmp(\"IBSUFpAliq\"))\n"
    b"    gIBSMunpAliq = IIf(IsNull(rEmp(\"IBSMunpAliq\")), 0, rEmp(\"IBSMunpAliq\"))\n"
    b"Else\n"
    b"    gCBSpAliq = 0: gIBSUFpAliq = 0: gIBSMunpAliq = 0\n"
    b"End If\n"
    b"If rEmp.State <> 0 Then rEmp.Close\n"
    b"Set rEmp = Nothing\n",
    b"",
    "carga_emp_cidade"
)

# 3. Adicionar cboTipoIS.AddItem
patch(
    b"    ' CST Imposto Seletivo\n"
    b"    cboISCST.AddItem \"00 - Incid\xeancia na Origem (Ind\xfastria/Importador)\"\n"
    b"    cboISCST.AddItem \"01 - Sa\xedda Tributada (Com\xe9rcio/Varejo)\"\n"
    b"    cboISCST.AddItem \"99 - Outras Opera\xe7\xf5es\"\n",
    b"    ' CST Imposto Seletivo\n"
    b"    cboISCST.AddItem \"00 - Incid\xeancia na Origem (Ind\xfastria/Importador)\"\n"
    b"    cboISCST.AddItem \"01 - Sa\xedda Tributada (Com\xe9rcio/Varejo)\"\n"
    b"    cboISCST.AddItem \"99 - Outras Opera\xe7\xf5es\"\n"
    b"    ' Tipo de calculo IS\n"
    b"    cboTipoIS.AddItem \"1 - Ad Valorem (%)\"\n"
    b"    cboTipoIS.AddItem \"2 - Ad Rem (R$/unid)\"\n"
    b"    cboTipoIS.AddItem \"3 - Misto (% + R$/unid)\"\n",
    "cboTipoIS_additem"
)

# 4. cboISClasse_Click: auto-preencher cboTipoIS
patch(
    b"Private Sub cboISClasse_Click()\n"
    b"    If cboISClasse.ListIndex < 0 Then lblISCSTClass.Caption = \"\": Exit Sub\n"
    b"    Dim sCod As String\n"
    b"    sCod = Left(cboISClasse.Text, 6)\n"
    b"    lblISCSTClass.Caption = Mid(cboISClasse.Text, 10)\n"
    b"    AplicarModoIS\n"
    b"End Sub\n",
    b"Private Sub cboISClasse_Click()\n"
    b"    If cboISClasse.ListIndex < 0 Then lblISCSTClass.Caption = \"\": Exit Sub\n"
    b"    Dim sCod As String\n"
    b"    Dim rTipo As ADODB.Recordset\n"
    b"    sCod = Left(cboISClasse.Text, 6)\n"
    b"    lblISCSTClass.Caption = Mid(cboISClasse.Text, 10)\n"
    b"    RsOpen rTipo, \"SELECT tipo_calculo_is FROM tbISClassTrib WHERE cClassTrib_IS = '\" & sCod & \"'\"\n"
    b"    If Not rTipo.EOF Then\n"
    b"        SelecionarNoCombo cboTipoIS, CStr(IIf(IsNull(rTipo(\"tipo_calculo_is\")), 1, rTipo(\"tipo_calculo_is\"))), True\n"
    b"    End If\n"
    b"    If rTipo.State <> 0 Then rTipo.Close\n"
    b"    AplicarModoIS\n"
    b"End Sub\n",
    "cboISClasse_Click"
)

# 5. Carregar cClassTrib_IS e tipo_calculo_is ao editar
patch(
    b"    SelecionarNoCombo cboISCST, ValidateNull(r(\"ISCST\")), True\n",
    b"    SelecionarNoCombo cboISCST, ValidateNull(r(\"ISCST\")), True\n"
    b"    SelecionarNoCombo cboISClasse, ValidateNull(r(\"cClassTrib_IS\")), True\n"
    b"    If Not IsNull(r(\"tipo_calculo_is\")) Then SelecionarNoCombo cboTipoIS, CStr(r(\"tipo_calculo_is\")), True\n",
    "carga_editar"
)

# 6. LimparObjetos: limpar cboISClasse e cboTipoIS
patch(
    b"SelecionarNoCombo cboISCST, \"99\", True\n"
    b"'cboCFOP.Text = \"\"\n",
    b"SelecionarNoCombo cboISCST, \"99\", True\n"
    b"cboISClasse.ListIndex = -1\n"
    b"cboTipoIS.ListIndex = -1\n"
    b"'cboCFOP.Text = \"\"\n",
    "limpar_combos_IS"
)

# 7. INSERT: coluna list
patch(
    b"\"categoria, PRATELEIRA, quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, ICMSCST, ICMSAliq, PISCST, COFINSCST, IPICST, NCM, CEST, CFOP, Alterado, PedirPeso, IPIALIQ, COFINSALIQ, PISALIQ, pRedBc, CODPROD_FRACAO, QUANT_FRACAO, modBC, modBCST, pMVAST, pRedBCST, pICMSST, IBSCBSCST, CBSpAliq, IBSUFpAliq, IBSMunpAliq, ISCST, ISpIS, EANEmbalagem, Fracionamento) VALUES (\" & _\n",
    b"\"categoria, PRATELEIRA, quant_min, INF_ADICIONA, quant_estoque, ref, tamanho, ICMSCST, ICMSAliq, PISCST, COFINSCST, IPICST, NCM, CEST, CFOP, Alterado, PedirPeso, IPIALIQ, COFINSALIQ, PISALIQ, pRedBc, CODPROD_FRACAO, QUANT_FRACAO, modBC, modBCST, pMVAST, pRedBCST, pICMSST, IBSCBSCST, ISCST, cClassTrib_IS, tipo_calculo_is, EANEmbalagem, Fracionamento) VALUES (\" & _\n",
    "insert_cols"
)

# 8. INSERT: linha de valores - parte IBS/IS
insert_old_vals = (
    b"   vModBC & \" , \" & vModBCST & \", \" & Replace(CDbl(txtMVA.Text), \",\", \".\") & \", \" "
    b"& Replace(CDbl(txtRedBCST.Text), \",\", \".\") & \", \" & Replace(CDbl(txtSTAliq.Text), \",\", \".\") "
    b"& \", '\" & Left(cboIBSCBSCST.Text, 3) & \"', 0, 0, 0, '\" & Left(cboISCST.Text, 2) "
    b"& \"', 0, '\" & txtEANCaixa.Text"
)
c = data.count(insert_old_vals)
if c == 1:
    insert_new_vals = (
        b"   vModBC & \" , \" & vModBCST & \", \" & Replace(CDbl(txtMVA.Text), \",\", \".\") & \", \""
        b" & Replace(CDbl(txtRedBCST.Text), \",\", \".\") & \", \" & Replace(CDbl(txtSTAliq.Text), \",\", \".\")"
        b" & \", '\" & Left(cboIBSCBSCST.Text, 3) & \"', '\" & Left(cboISCST.Text, 2)"
        b" & \"', \" & IIf(cboISClasse.ListIndex >= 0, \"'\" & Left(cboISClasse.Text, 6) & \"'\", \"NULL\")"
        b" & \", \" & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\")"
        b" & \", '\" & txtEANCaixa.Text"
    )
    data = data.replace(insert_old_vals, insert_new_vals, 1)
else:
    print(f"ERRO [insert_vals] ({c}x)")
    errors += 1
    # debug
    idx = data.find(b"0, 0, 0, '\" & Left(cboISCST")
    if idx >= 0:
        print(repr(data[idx-50:idx+150]))

# 9. UPDATE: bloco reforma
patch(
    b"    ' --- Bloco 3: Reforma Tribut\xc1RIA ---\n"
    b"    sSQL = sSQL & \"IBSCBSCST = '\" & Left(cboIBSCBSCST.Text, 3) & \"', \"\n"
    b"    sSQL = sSQL & \"CBSpAliq = 0, \"\n"
    b"    sSQL = sSQL & \"IBSUFpAliq = 0, \"\n"
    b"    sSQL = sSQL & \"IBSMunpAliq = 0, \"\n"
    b"    sSQL = sSQL & \"ISCST = '\" & Left(cboISCST.Text, 2) & \"', \"\n"
    b"    sSQL = sSQL & \"ISpIS = 0\"\n",
    b"    ' --- Bloco 3: Reforma Tribut\xc1RIA ---\n"
    b"    sSQL = sSQL & \"IBSCBSCST = '\" & Left(cboIBSCBSCST.Text, 3) & \"', \"\n"
    b"    sSQL = sSQL & \"ISCST = '\" & Left(cboISCST.Text, 2) & \"', \"\n"
    b"    sSQL = sSQL & \"cClassTrib_IS = \" & IIf(cboISClasse.ListIndex >= 0, \"'\" & Left(cboISClasse.Text, 6) & \"'\", \"NULL\") & \", \"\n"
    b"    sSQL = sSQL & \"tipo_calculo_is = \" & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\")\n",
    "update_reforma"
)

# 10. PreencherReformaDoNCM: selecionar cboTipoIS
patch(
    b"    If Len(sClassIS) = 6 And iTipoIS > 0 Then\n"
    b"        SelecionarNoCombo cboISCST, \"01\", True\n"
    b"        SelecionarNoCombo cboISClasse, sClassIS, True\n"
    b"    End If\n"
    b"End Sub\n",
    b"    If Len(sClassIS) = 6 And iTipoIS > 0 Then\n"
    b"        SelecionarNoCombo cboISCST, \"01\", True\n"
    b"        SelecionarNoCombo cboISClasse, sClassIS, True\n"
    b"        SelecionarNoCombo cboTipoIS, CStr(iTipoIS), True\n"
    b"    End If\n"
    b"End Sub\n",
    "PreencherReformaDoNCM"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
