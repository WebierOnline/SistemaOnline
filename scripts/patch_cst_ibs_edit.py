"""
patch_cst_ibs_edit.py
Aplica 4 mudanças em NFe_Completa.frm:
  P1 - FormatarGridItensNota: move col 9 do array azure para amarelo (editavel)
  P2 - GridNotasItens_Click: adiciona col 9 ao Case editável
  P3 - txtEdit_LostFocus Dims: adiciona variáveis para recálculo IBS/CBS
  P4 - txtEdit_LostFocus Cases 8/9/10: insere Case 9 (CST IBS) e atualiza Case 10 (CLASS IBS)
"""
import sys, shutil, re

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cst_ibs_edit"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    raw = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

data = raw

patches = []

# ── P1: FormatarGridItensNota — arrays de cor ────────────────────────────────
# Remove col 9 do azure e adiciona ao amarelo
patches.append((
    b"      For Each colEdit In Array(2, 5, 6, 7, 8, 10, 22, 24, 26, 28, 29, 30)",
    b"      For Each colEdit In Array(2, 5, 6, 7, 8, 9, 10, 22, 24, 26, 28, 29, 30)"
))
patches.append((
    b"      For Each colRef In Array(9, 11, 12, 13)",
    b"      For Each colRef In Array(11, 12, 13)"
))

# ── P2: GridNotasItens_Click — adiciona col 9 ao Case ────────────────────────
patches.append((
    b"    Case 2, 5, 6, 7, 8, 10, 22, 24, 26, 28, 29, 30",
    b"    Case 2, 5, 6, 7, 8, 9, 10, 22, 24, 26, 28, 29, 30"
))

# ── P3: txtEdit_LostFocus — adiciona Dims para IBS/CBS ───────────────────────
patches.append((
    b"Dim dblMVA    As Double\r\n",
    b"Dim dblMVA    As Double\r\n"
    b"Dim dblPRedIBS As Double\r\n"
    b"Dim dblPRedCBS As Double\r\n"
    b"Dim dblIBSvBC As Double\r\n"
    b"Dim dblIBSUFpAliq As Double\r\n"
    b"Dim dblIBSMunpAliq As Double\r\n"
    b"Dim dblCBSvBC As Double\r\n"
    b"Dim dblCBSpAliq As Double\r\n"
    b"Dim curIBSvIBSUF As Currency\r\n"
    b"Dim curIBSvIBSMun As Currency\r\n"
    b"Dim curIBSvIBS As Currency\r\n"
    b"Dim curCBSvCBS As Currency\r\n"
    b"Dim sChkCST As String\r\n"
    b"Dim sNewClassTrib As String\r\n"
    b"Dim rsClassTrib As ADODB.Recordset\r\n"
    b"Dim rsIBSItem As ADODB.Recordset\r\n"
))

# ── P4: Case 8 + antigo Case 10 → Case 8 + Case 9 (novo) + Case 10 atualizado
OLD_CASES = (
    b"    Case 8 ' CST\r\n"
    b"        If sVal = \"\" Or Len(sVal) <> 3 Then\r\n"
    b"            MsgBox \"CST deve ter 3 d\xedgitos!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET CST = '\" & sVal & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"\r\n"
    b"    Case 10 ' cClassTrib\r\n"
    b"        sVal = UCase(Trim(sVal))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib = '\" & Replace(sVal, \"'\", \"''\") & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
)

NEW_CASES = (
    b"    Case 8 ' CST\r\n"
    b"        If sVal = \"\" Or Len(sVal) <> 3 Then\r\n"
    b"            MsgBox \"CST deve ter 3 d\xedgitos!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET CST = '\" & sVal & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"\r\n"
    b"    Case 9 ' CST IBS\r\n"
    b"        sVal = Trim(sVal)\r\n"
    b"        If sVal = \"\" Then\r\n"
    b"            MsgBox \"CST IBS n\xe3o pode ser vazio!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        sChkCST = SQLExecutaRetorno(\"SELECT TOP 1 CST FROM TbIBSCBSClassTrib WHERE CST = '\" & Replace(sVal, \"'\", \"''\") & \"'\", \"CST\", \"\")\r\n"
    b"        If sChkCST = \"\" Then\r\n"
    b"            MsgBox \"CST IBS '\" & sVal & \"' n\xe3o encontrado em TbIBSCBSClassTrib!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        Set rsClassTrib = New ADODB.Recordset\r\n"
    b"        RsOpen rsClassTrib, \"SELECT TOP 1 cClassTrib, pRedIBS, pRedCBS FROM TbIBSCBSClassTrib WHERE CST = '\" & Replace(sVal, \"'\", \"''\") & \"'\"\r\n"
    b"        If rsClassTrib.EOF Then\r\n"
    b"            rsClassTrib.Close: Set rsClassTrib = Nothing\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        sNewClassTrib = rsClassTrib!cClassTrib & \"\"\r\n"
    b"        dblPRedIBS = CDbl(rsClassTrib!pRedIBS)\r\n"
    b"        dblPRedCBS = CDbl(rsClassTrib!pRedCBS)\r\n"
    b"        rsClassTrib.Close: Set rsClassTrib = Nothing\r\n"
    b"        Set rsIBSItem = New ADODB.Recordset\r\n"
    b"        RsOpen rsIBSItem, \"SELECT IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, CBS_vBC, CBS_pAliq FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        If rsIBSItem.EOF Then\r\n"
    b"            rsIBSItem.Close: Set rsIBSItem = Nothing\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblIBSvBC = CDbl(rsIBSItem!IBS_vBC)\r\n"
    b"        dblIBSUFpAliq = CDbl(rsIBSItem!IBS_UFpAliq)\r\n"
    b"        dblIBSMunpAliq = CDbl(rsIBSItem!IBS_MunpAliq)\r\n"
    b"        dblCBSvBC = CDbl(rsIBSItem!CBS_vBC)\r\n"
    b"        dblCBSpAliq = CDbl(rsIBSItem!CBS_pAliq)\r\n"
    b"        rsIBSItem.Close: Set rsIBSItem = Nothing\r\n"
    b"        curIBSvIBSUF = CCur(Format(dblIBSvBC * dblIBSUFpAliq * (1 - dblPRedIBS / 100) / 100, \"0.00\"))\r\n"
    b"        curIBSvIBSMun = CCur(Format(dblIBSvBC * dblIBSMunpAliq * (1 - dblPRedIBS / 100) / 100, \"0.00\"))\r\n"
    b"        curIBSvIBS = curIBSvIBSUF + curIBSvIBSMun\r\n"
    b"        curCBSvCBS = CCur(Format(dblCBSvBC * dblCBSpAliq * (1 - dblPRedCBS / 100) / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IBSCBS_CST = '\" & Replace(sVal, \"'\", \"''\") & \"', cClassTrib = '\" & Replace(sNewClassTrib, \"'\", \"''\") & \"', IBS_pRed = \" & FSQL(dblPRedIBS, 4) & \", CBS_pRed = \" & FSQL(dblPRedCBS, 4) & \", IBS_vIBSUF = \" & FSQL(curIBSvIBSUF, 2) & \", IBS_vIBSMun = \" & FSQL(curIBSvIBSMun, 2) & \", IBS_vIBS = \" & FSQL(curIBSvIBS, 2) & \", CBS_vCBS = \" & FSQL(curCBSvCBS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 9) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 10) = sNewClassTrib\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 11) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 12) = FormatNumber(curCBSvCBS, 2)\r\n"
    b"\r\n"
    b"    Case 10 ' cClassTrib (CLASS IBS)\r\n"
    b"        sVal = UCase(Trim(sVal))\r\n"
    b"        If sVal = \"\" Then\r\n"
    b"            MsgBox \"CLASS IBS n\xe3o pode ser vazio!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        Dim sCSTAtual As String\r\n"
    b"        sCSTAtual = Trim(GridNotasItens.TextMatrix(iRow, 9))\r\n"
    b"        sChkCST = SQLExecutaRetorno(\"SELECT TOP 1 cClassTrib FROM TbIBSCBSClassTrib WHERE CST = '\" & Replace(sCSTAtual, \"'\", \"''\") & \"' AND cClassTrib = '\" & Replace(sVal, \"'\", \"''\") & \"'\", \"cClassTrib\", \"\")\r\n"
    b"        If sChkCST = \"\" Then\r\n"
    b"            MsgBox \"CLASS IBS '\" & sVal & \"' n\xe3o pertence ao CST '\" & sCSTAtual & \"' em TbIBSCBSClassTrib!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        Set rsClassTrib = New ADODB.Recordset\r\n"
    b"        RsOpen rsClassTrib, \"SELECT pRedIBS, pRedCBS FROM TbIBSCBSClassTrib WHERE CST = '\" & Replace(sCSTAtual, \"'\", \"''\") & \"' AND cClassTrib = '\" & Replace(sVal, \"'\", \"''\") & \"'\"\r\n"
    b"        If rsClassTrib.EOF Then\r\n"
    b"            rsClassTrib.Close: Set rsClassTrib = Nothing\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblPRedIBS = CDbl(rsClassTrib!pRedIBS)\r\n"
    b"        dblPRedCBS = CDbl(rsClassTrib!pRedCBS)\r\n"
    b"        rsClassTrib.Close: Set rsClassTrib = Nothing\r\n"
    b"        Set rsIBSItem = New ADODB.Recordset\r\n"
    b"        RsOpen rsIBSItem, \"SELECT IBS_vBC, IBS_UFpAliq, IBS_MunpAliq, CBS_vBC, CBS_pAliq FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        If rsIBSItem.EOF Then\r\n"
    b"            rsIBSItem.Close: Set rsIBSItem = Nothing\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dblIBSvBC = CDbl(rsIBSItem!IBS_vBC)\r\n"
    b"        dblIBSUFpAliq = CDbl(rsIBSItem!IBS_UFpAliq)\r\n"
    b"        dblIBSMunpAliq = CDbl(rsIBSItem!IBS_MunpAliq)\r\n"
    b"        dblCBSvBC = CDbl(rsIBSItem!CBS_vBC)\r\n"
    b"        dblCBSpAliq = CDbl(rsIBSItem!CBS_pAliq)\r\n"
    b"        rsIBSItem.Close: Set rsIBSItem = Nothing\r\n"
    b"        curIBSvIBSUF = CCur(Format(dblIBSvBC * dblIBSUFpAliq * (1 - dblPRedIBS / 100) / 100, \"0.00\"))\r\n"
    b"        curIBSvIBSMun = CCur(Format(dblIBSvBC * dblIBSMunpAliq * (1 - dblPRedIBS / 100) / 100, \"0.00\"))\r\n"
    b"        curIBSvIBS = curIBSvIBSUF + curIBSvIBSMun\r\n"
    b"        curCBSvCBS = CCur(Format(dblCBSvBC * dblCBSpAliq * (1 - dblPRedCBS / 100) / 100, \"0.00\"))\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib = '\" & Replace(sVal, \"'\", \"''\") & \"', IBS_pRed = \" & FSQL(dblPRedIBS, 4) & \", CBS_pRed = \" & FSQL(dblPRedCBS, 4) & \", IBS_vIBSUF = \" & FSQL(curIBSvIBSUF, 2) & \", IBS_vIBSMun = \" & FSQL(curIBSvIBSMun, 2) & \", IBS_vIBS = \" & FSQL(curIBSvIBS, 2) & \", CBS_vCBS = \" & FSQL(curCBSvCBS, 2) & \" WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 11) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 12) = FormatNumber(curCBSvCBS, 2)\r\n"
)

patches.append((OLD_CASES, NEW_CASES))

# ── Aplicar patches ────────────────────────────────────────────────────────────
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
