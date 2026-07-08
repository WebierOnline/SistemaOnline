"""
Produtos_Cadastro.frm:
1. AtualizarReforma: remove IBS/CBS/IS combos, mantém só PIS/COFINS
2. PreencherReformaDoNCM: busca CST do IBS via TbIBSCBSClassTrib e CST do IS via tbISClassTrib
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
        print(f"OK: {label}")

# ── 1. AtualizarReforma: remove Dim sIBSCST/sCClassTrib ──────────────────────
patch(
    b"    Dim sIBSCST As String\n"
    b"    Dim sCClassTrib As String\n"
    b"    sCat = UCase(cboCategoria.Text)\n"
    b"\n"
    b"    ' Reset inicial\n"
    b"    SelecionarNoCombo cboISCST, \"99\", True\n",
    b"    sCat = UCase(cboCategoria.Text)\n",
    "AtualizarReforma Dims"
)

# ── 2. Remove sIBSCST/sCClassTrib de cada Case ────────────────────────────────
# Cesta Basica/Hortifrutti/Carnes
patch(
    b"            sIBSCST = \"400\": sCClassTrib = \"400001\"\n"
    b"            sPISCST = \"06\": sCOFINSCST = \"06\"\n",
    b"            sPISCST = \"06\": sCOFINSCST = \"06\"\n",
    "Case 400"
)
# Limpeza/Higiene/Bebidas
patch(
    b"            sIBSCST = \"200\": sCClassTrib = \"200001\"\n"
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\n",
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\n",
    "Case 200"
)
# Bebidas alco/acucaradas/tabacaria
patch(
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\n"
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\n"
    b"            SelecionarNoCombo cboISCST, \"01\", True\n",
    b"            sPISCST = \"04\": sCOFINSCST = \"04\"\n"
    b"            dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\n",
    "Case 000 bebidas"
)
# Alimentos/GERAL/Frios (1a ocorrencia = Case padrao)
patch(
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\n"
    b"            If var_CRT = 3 Then\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\n"
    b"            Else\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\n"
    b"            End If\n"
    b"\n"
    b"        ' --- Case Else: CST 000 gen\xe9rico ---\n"
    b"        Case Else\n"
    b"            sIBSCST = \"000\": sCClassTrib = \"000001\"\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\n",
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\n"
    b"            If var_CRT = 3 Then\n"
    b"                dPISAliq = \"0,65\": dCOFINSAliq = \"3,00\"\n"
    b"            Else\n"
    b"                dPISAliq = \"0,00\": dCOFINSAliq = \"0,00\"\n"
    b"            End If\n"
    b"\n"
    b"        ' --- Case Else: gen\xe9rico ---\n"
    b"        Case Else\n"
    b"            sPISCST = \"01\": sCOFINSCST = \"01\"\n",
    "Case 000 alimentos+else"
)

# ── 3. Remove SelecionarNoCombo cboIBSCBSCST/cboIBSCBSClasse e comentários ───
patch(
    b"    ' Seleciona CST (dispara cboIBSCBSCST_Click que preenche cboIBSCBSClasse)\n"
    b"    SelecionarNoCombo cboIBSCBSCST, sIBSCST, True\n"
    b"    ' Seleciona a classificacao especifica (dispara cboIBSCBSClasse_Click com reducao)\n"
    b"    SelecionarNoCombo cboIBSCBSClasse, sCClassTrib\n"
    b"    txtPISCST.Text = sPISCST\n",
    b"    txtPISCST.Text = sPISCST\n",
    "AtualizarReforma remove IBS selects"
)

# ── 4. PreencherReformaDoNCM: reescrever para buscar CST do banco ─────────────
patch(
    b"Private Sub PreencherReformaDoNCM()\n"
    b"    Dim rNCM As ADODB.Recordset\n"
    b"    Dim sClassIBS As String\n"
    b"    Dim sClassIS As String\n"
    b"    Dim iTipoIS As Integer\n"
    b"    If Len(Trim(txtNCM.Text)) <> 8 Then Exit Sub\n"
    b"    RsOpen rNCM, \"SELECT cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is FROM tbNCM WHERE NCM='\" & Trim(txtNCM.Text) & \"'\"\n"
    b"    If rNCM.EOF Then\n"
    b"        If rNCM.State <> 0 Then rNCM.Close\n"
    b"        Exit Sub\n"
    b"    End If\n"
    b"    sClassIBS = \"\"\n"
    b"    If Not IsNull(rNCM(\"cClassTrib_IBS\")) Then sClassIBS = Trim(CStr(rNCM(\"cClassTrib_IBS\")))\n"
    b"    sClassIS = \"\"\n"
    b"    If Not IsNull(rNCM(\"cClassTrib_IS\")) Then sClassIS = Trim(CStr(rNCM(\"cClassTrib_IS\")))\n"
    b"    iTipoIS = 0\n"
    b"    If Not IsNull(rNCM(\"tipo_calculo_is\")) Then iTipoIS = CInt(rNCM(\"tipo_calculo_is\"))\n"
    b"    If rNCM.State <> 0 Then rNCM.Close\n"
    b"    ' IBS/CBS: CST pelos 3 primeiros chars, classe pelo codigo completo\n"
    b"    If Len(sClassIBS) >= 3 Then\n"
    b"        SelecionarNoCombo cboIBSCBSCST, Left(sClassIBS, 3), True\n"
    b"        If Len(sClassIBS) = 6 Then SelecionarNoCombo cboIBSCBSClasse, sClassIBS\n"
    b"    End If\n"
    b"    ' IS: se tem classe e tipo de calculo, marca CST 01 e seleciona classe\n"
    b"    If Len(sClassIS) = 6 And iTipoIS > 0 Then\n"
    b"        SelecionarNoCombo cboISCST, \"01\", True\n"
    b"        SelecionarNoCombo cboISClasse, sClassIS, True\n"
    b"        SelecionarNoCombo cboTipoIS, CStr(iTipoIS), True\n"
    b"    End If\n"
    b"End Sub\n",
    b"Private Sub PreencherReformaDoNCM()\n"
    b"    Dim rNCM As ADODB.Recordset\n"
    b"    Dim rCST As ADODB.Recordset\n"
    b"    Dim sClassIBS As String, sClassIS As String\n"
    b"    Dim iTipoIS As Integer\n"
    b"    Dim sIBSCST As String, sISCST As String\n"
    b"    If Len(Trim(txtNCM.Text)) <> 8 Then Exit Sub\n"
    b"    RsOpen rNCM, \"SELECT cClassTrib_IBS, cClassTrib_IS, tipo_calculo_is FROM tbNCM WHERE NCM='\" & Trim(txtNCM.Text) & \"'\"\n"
    b"    If rNCM.EOF Then\n"
    b"        If rNCM.State <> 0 Then rNCM.Close\n"
    b"        Exit Sub\n"
    b"    End If\n"
    b"    sClassIBS = \"\"\n"
    b"    If Not IsNull(rNCM(\"cClassTrib_IBS\")) Then sClassIBS = Trim(CStr(rNCM(\"cClassTrib_IBS\")))\n"
    b"    sClassIS = \"\"\n"
    b"    If Not IsNull(rNCM(\"cClassTrib_IS\")) Then sClassIS = Trim(CStr(rNCM(\"cClassTrib_IS\")))\n"
    b"    iTipoIS = 0\n"
    b"    If Not IsNull(rNCM(\"tipo_calculo_is\")) Then iTipoIS = CInt(rNCM(\"tipo_calculo_is\"))\n"
    b"    If rNCM.State <> 0 Then rNCM.Close\n"
    b"    ' IBS/CBS: busca CST pelo cClassTrib na TbIBSCBSClassTrib\n"
    b"    If Len(sClassIBS) = 6 Then\n"
    b"        RsOpen rCST, \"SELECT CST FROM TbIBSCBSClassTrib WHERE cClassTrib = '\" & sClassIBS & \"'\"\n"
    b"        If Not rCST.EOF Then\n"
    b"            sIBSCST = Trim(CStr(rCST(\"CST\")))\n"
    b"            SelecionarNoCombo cboIBSCBSCST, sIBSCST, True\n"
    b"            SelecionarNoCombo cboIBSCBSClasse, sClassIBS\n"
    b"        End If\n"
    b"        If rCST.State <> 0 Then rCST.Close\n"
    b"    End If\n"
    b"    ' IS: busca ISCST pelo cClassTrib_IS na tbISClassTrib\n"
    b"    If Len(sClassIS) = 6 And iTipoIS > 0 Then\n"
    b"        RsOpen rCST, \"SELECT ISNULL(ISCST,'01') AS ISCST FROM tbISClassTrib WHERE cClassTrib_IS = '\" & sClassIS & \"'\"\n"
    b"        If Not rCST.EOF Then\n"
    b"            sISCST = Trim(CStr(rCST(\"ISCST\")))\n"
    b"            SelecionarNoCombo cboISCST, sISCST, True\n"
    b"            SelecionarNoCombo cboISClasse, sClassIS, True\n"
    b"            SelecionarNoCombo cboTipoIS, CStr(iTipoIS), True\n"
    b"        End If\n"
    b"        If rCST.State <> 0 Then rCST.Close\n"
    b"    End If\n"
    b"End Sub\n",
    "PreencherReformaDoNCM"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
