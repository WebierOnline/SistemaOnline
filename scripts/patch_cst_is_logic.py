"""
patch_cst_is_logic.py
Reescreve Case 11 (CST IS) e Case 12 (CLASS IS) em txtEdit_LostFocus com:
  - CST IS aceita somente vazio / '00' / '01' / '99'
  - vazio/00/99 zera CLASS IS e todos os campos IS do item
  - CLASS IS valida existência em tbISClassTrib.cClassTrib_IS (sem filtro por ISCST)
  - CLASS IS válido: busca dados (vigência), recalcula IS_vIS, atualiza DB + grid
  - Adiciona Dims necessários ao topo do Sub
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cst_is_logic"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ── P1: Novos Dims após rsISTrib ─────────────────────────────────────────────
patches.append((
    b"Dim rsISTrib As ADODB.Recordset\r\n",
    b"Dim rsISTrib As ADODB.Recordset\r\n"
    b"Dim sISCST2 As String\r\n"
    b"Dim iTipoIS2 As Integer\r\n"
    b"Dim dISpAliq2 As Double\r\n"
    b"Dim dISvUnid2 As Double\r\n"
    b"Dim sISTrib2uTrib As String\r\n"
    b"Dim dISvBC2 As Double\r\n"
    b"Dim dISqUnid2 As Double\r\n"
    b"Dim dCalcISvBC2 As Double\r\n"
    b"Dim curISvIS2 As Currency\r\n"
))

# ── P2: Substituir Cases 11 e 12 completos ───────────────────────────────────
OLD_CASES = (
    b"    Case 11 ' CST IS\r\n"
    b"        sVal = Trim(sVal)\r\n"
    b"        If sVal = \"\" Then\r\n"
    b"            MsgBox \"CST IS n\xe3o pode ser vazio!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        sChkCST = SQLExecutaRetorno(\"SELECT TOP 1 ISCST FROM tbISClassTrib WHERE ISCST = '\" & Replace(sVal, \"'\", \"''\") & \"'\", \"ISCST\", \"\")\r\n"
    b"        If sChkCST = \"\" Then\r\n"
    b"            MsgBox \"CST IS '\" & sVal & \"' n\xe3o encontrado em tbISClassTrib!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        Set rsISTrib = New ADODB.Recordset\r\n"
    b"        RsOpen rsISTrib, \"SELECT TOP 1 cClassTrib_IS FROM tbISClassTrib WHERE ISCST = '\" & Replace(sVal, \"'\", \"''\") & \"'\"\r\n"
    b"        If rsISTrib.EOF Then\r\n"
    b"            rsISTrib.Close: Set rsISTrib = Nothing\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        sNewClassIS = rsISTrib!cClassTrib_IS & \"\"\r\n"
    b"        rsISTrib.Close: Set rsISTrib = Nothing\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '\" & Replace(sVal, \"'\", \"''\") & \"', cClassTrib_IS = '\" & Replace(sNewClassIS, \"'\", \"''\") & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 11) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 12) = sNewClassIS\r\n"
    b"\r\n"
    b"    Case 12 ' cClassTrib_IS (CLASS IS)\r\n"
    b"        sVal = UCase(Trim(sVal))\r\n"
    b"        If sVal = \"\" Then\r\n"
    b"            MsgBox \"CLASS IS n\xe3o pode ser vazio!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        sISCSTAtual = Trim(GridNotasItens.TextMatrix(iRow, 11))\r\n"
    b"        sChkCST = SQLExecutaRetorno(\"SELECT TOP 1 cClassTrib_IS FROM tbISClassTrib WHERE ISCST = '\" & Replace(sISCSTAtual, \"'\", \"''\") & \"' AND cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"'\", \"cClassTrib_IS\", \"\")\r\n"
    b"        If sChkCST = \"\" Then\r\n"
    b"            MsgBox \"CLASS IS '\" & sVal & \"' n\xe3o pertence ao CST '\" & sISCSTAtual & \"' em tbISClassTrib!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
)

NEW_CASES = (
    b"    Case 11 ' CST IS\r\n"
    b"        sVal = Trim(sVal)\r\n"
    b"        ' Aceitar somente: vazio, 00, 01, 99\r\n"
    b"        If sVal <> \"\" And sVal <> \"00\" And sVal <> \"01\" And sVal <> \"99\" Then\r\n"
    b"            MsgBox \"CST IS inv\xe1lido! Aceitos: vazio, 00, 01, 99\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        If sVal = \"\" Or sVal = \"00\" Or sVal = \"99\" Then\r\n"
    b"            ' Sem IS: zerar todos os campos IS do item\r\n"
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '\" & sVal & \"', cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 11) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"        Else\r\n"
    b"            ' CST = 01: apenas atualiza IS_CST, mant\xe9m CLASS IS\r\n"
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '01' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 11) = \"01\"\r\n"
    b"        End If\r\n"
    b"\r\n"
    b"    Case 12 ' cClassTrib_IS (CLASS IS)\r\n"
    b"        sVal = Trim(sVal)\r\n"
    b"        sISCSTAtual = Trim(GridNotasItens.TextMatrix(iRow, 11))\r\n"
    b"        ' CLASS IS s\xf3 faz sentido com CST IS = 01\r\n"
    b"        If sISCSTAtual <> \"01\" Then\r\n"
    b"            MsgBox \"CLASS IS s\xf3 pode ser preenchido quando CST IS = '01'!\", vbInformation, \"Aviso\"\r\n"
    b"            Exit Sub\r\n"
    b"        End If\r\n"
    b"        If sVal = \"\" Then\r\n"
    b"            ' Limpar CLASS IS: zerar campos IS\r\n"
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, iCol) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"        Else\r\n"
    b"            ' Validar exist\xeancia em tbISClassTrib\r\n"
    b"            sChkCST = SQLExecutaRetorno(\"SELECT TOP 1 cClassTrib_IS FROM tbISClassTrib WHERE cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"'\", \"cClassTrib_IS\", \"\")\r\n"
    b"            If sChkCST = \"\" Then\r\n"
    b"                MsgBox \"CLASS IS '\" & sVal & \"' n\xe3o encontrado em tbISClassTrib!\", vbInformation, \"Aviso\"\r\n"
    b"                Exit Sub\r\n"
    b"            End If\r\n"
    b"            ' Buscar dados IS com prioridade de vig\xeancia\r\n"
    b"            Set rsISTrib = New ADODB.Recordset\r\n"
    b"            RsOpen rsISTrib, \"SELECT TOP 1 ISCST, tipo_calculo_is, ISpAliq, ISvUnid, uTrib_IS \" & _\r\n"
    b"                             \"FROM tbISClassTrib \" & _\r\n"
    b"                             \"WHERE cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"' \" & _\r\n"
    b"                             \"ORDER BY CASE WHEN GETDATE() BETWEEN dIniVig AND dFimVig THEN 0 ELSE 1 END ASC, dFimVig DESC\"\r\n"
    b"            If rsISTrib.EOF Then\r\n"
    b"                rsISTrib.Close: Set rsISTrib = Nothing\r\n"
    b"                Exit Sub\r\n"
    b"            End If\r\n"
    b"            sISCST2 = IIf(IsNull(rsISTrib!ISCST), \"\", CStr(rsISTrib!ISCST))\r\n"
    b"            iTipoIS2 = CInt(IIf(IsNull(rsISTrib!tipo_calculo_is), 0, rsISTrib!tipo_calculo_is))\r\n"
    b"            dISpAliq2 = CDbl(IIf(IsNull(rsISTrib!ISpAliq), 0, rsISTrib!ISpAliq))\r\n"
    b"            dISvUnid2 = CDbl(IIf(IsNull(rsISTrib!ISvUnid), 0, rsISTrib!ISvUnid))\r\n"
    b"            sISTrib2uTrib = IIf(IsNull(rsISTrib!uTrib_IS), \"\", CStr(rsISTrib!uTrib_IS))\r\n"
    b"            rsISTrib.Close: Set rsISTrib = Nothing\r\n"
    b"            ' Ler IS_vBC e IS_qUnid atuais para recalcular\r\n"
    b"            Set rsIBSItem = New ADODB.Recordset\r\n"
    b"            RsOpen rsIBSItem, \"SELECT IS_vBC, IS_qUnid FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            dISvBC2 = 0: dISqUnid2 = 0\r\n"
    b"            If Not rsIBSItem.EOF Then\r\n"
    b"                dISvBC2 = CDbl(IIf(IsNull(rsIBSItem!IS_vBC), 0, rsIBSItem!IS_vBC))\r\n"
    b"                dISqUnid2 = CDbl(IIf(IsNull(rsIBSItem!IS_qUnid), 0, rsIBSItem!IS_qUnid))\r\n"
    b"            End If\r\n"
    b"            rsIBSItem.Close: Set rsIBSItem = Nothing\r\n"
    b"            ' IS_vBC s\xf3 se aplica nos tipos ad valorem (1 e 3)\r\n"
    b"            dCalcISvBC2 = IIf(iTipoIS2 = 1 Or iTipoIS2 = 3, dISvBC2, 0)\r\n"
    b"            ' qUnid e vUnid zerados nos tipos sem quantidade (0 e 1)\r\n"
    b"            If iTipoIS2 = 0 Or iTipoIS2 = 1 Then dISqUnid2 = 0: dISvUnid2 = 0\r\n"
    b"            ' Calcular IS_vIS por tipo\r\n"
    b"            Select Case iTipoIS2\r\n"
    b"                Case 1: curISvIS2 = CCur(Format(dCalcISvBC2 * dISpAliq2 / 100, \"0.00\"))\r\n"
    b"                Case 2: curISvIS2 = CCur(Format(dISqUnid2 * dISvUnid2, \"0.00\"))\r\n"
    b"                Case 3: curISvIS2 = CCur(Format(dCalcISvBC2 * dISpAliq2 / 100 + dISqUnid2 * dISvUnid2, \"0.00\"))\r\n"
    b"                Case Else: curISvIS2 = 0\r\n"
    b"            End Select\r\n"
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '\" & Replace(sISCST2, \"'\", \"''\") & \"', cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"', IS_tipo_calculo = \" & iTipoIS2 & \", IS_vBC = \" & FSQL(dCalcISvBC2, 2) & \", IS_pAliq = \" & FSQL(dISpAliq2, 4) & \", IS_qUnid = \" & FSQL(dISqUnid2, 4) & \", IS_vUnid = \" & FSQL(dISvUnid2, 4) & \", IS_vIS = \" & FSQL(curISvIS2, 2) & \", uTrib_IS = '\" & Replace(sISTrib2uTrib, \"'\", \"''\") & \"' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 11) = sISCST2\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(curISvIS2, 2)\r\n"
    b"        End If\r\n"
)

patches.append((OLD_CASES, NEW_CASES))

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
