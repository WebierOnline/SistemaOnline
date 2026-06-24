"""
patch_cst_is_edit.py
Torna CST IS (col 11) e CLASS IS (col 12) editáveis no GridNotasItens,
com validação contra tbISClassTrib.ISCST / cClassTrib_IS.
Mesmo padrão de CST IBS / CLASS IBS.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cst_is_edit"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ── P1: GridNotasItens_Click — adiciona cols 11 e 12 ao Case ─────────────────
patches.append((
    b"    Case 2, 5, 6, 7, 8, 9, 10, 24, 26, 28, 30, 31, 32",
    b"    Case 2, 5, 6, 7, 8, 9, 10, 11, 12, 24, 26, 28, 30, 31, 32"
))

# ── P2: txtEdit_LostFocus — novos Dims para IS ───────────────────────────────
patches.append((
    b"Dim rsIBSItem As ADODB.Recordset\r\n",
    b"Dim rsIBSItem As ADODB.Recordset\r\n"
    b"Dim sNewClassIS As String\r\n"
    b"Dim sISCSTAtual As String\r\n"
    b"Dim rsISTrib As ADODB.Recordset\r\n"
))

# ── P3: txtEdit_LostFocus — inserir Case 11 e Case 12 após Case 10 ────────────
patches.append((
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
    b"\r\n"
    b"    Case 24 ' %ICMS\r\n",

    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
    b"\r\n"
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
    b"\r\n"
    b"    Case 24 ' %ICMS\r\n"
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
