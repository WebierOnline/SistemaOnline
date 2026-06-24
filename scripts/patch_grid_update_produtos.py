"""
patch_grid_update_produtos.py

Ao editar CST IBS (col9), CLASS IBS (col10), CST IS (col11) e CLASS IS (col12)
no GridNotasItens, apos validar e salvar em NotaFiscalItens, propaga a mudanca
para a tabela produtos (CodigoProduto esta na col3 do grid).

Campos atualizados em produtos:
  col9  -> IBSCBSCST + cClassTrib (cClassTrib vem do lookup em TbIBSCBSClassTrib)
  col10 -> cClassTrib
  col11 (sem IS) -> ISCST = sVal, cClassTrib_IS = NULL, tipo_calculo_is = 1
  col11 (CST=01) -> ISCST = '01'
  col12 (limpar)  -> cClassTrib_IS = NULL, tipo_calculo_is = 1
  col12 (valido)  -> ISCST = sISCST2, cClassTrib_IS = sVal, tipo_calculo_is = iTipoIS2
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_grid_update_produtos"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# P1 - Case 9 (CST IBS): apos atualizar grid cols 9,10,13,14 -> propaga para produtos
patches.append((
    b"        GridNotasItens.TextMatrix(iRow, 9) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 10) = sNewClassTrib\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n",

    b"        GridNotasItens.TextMatrix(iRow, 9) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 10) = sNewClassTrib\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
    b"        dbData.Execute \"UPDATE produtos SET IBSCBSCST = '\" & Replace(sVal, \"'\", \"''\") & \"', cClassTrib = '\" & Replace(sNewClassTrib, \"'\", \"''\") & \"' WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n",

    "P1: Case9 CST IBS -> produtos"
))

# P2 - Case 10 (CLASS IBS): apos atualizar grid iCol,13,14 -> propaga cClassTrib para produtos
patches.append((
    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n",

    b"        GridNotasItens.TextMatrix(iRow, iCol) = sVal\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 13) = FormatNumber(curIBSvIBS, 2)\r\n"
    b"        GridNotasItens.TextMatrix(iRow, 14) = FormatNumber(curCBSvCBS, 2)\r\n"
    b"        dbData.Execute \"UPDATE produtos SET cClassTrib = '\" & Replace(sVal, \"'\", \"''\") & \"' WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n",

    "P2: Case10 CLASS IBS -> produtos"
))

# P3a - Case 11 branch "sem IS" (sVal = ""/00/99):
#   apos cols 11/12/15, antes do Else -> propaga ISCST+cClassTrib_IS+tipo_calculo_is
patches.append((
    b"            GridNotasItens.TextMatrix(iRow, 11) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"        Else\r\n",

    b"            GridNotasItens.TextMatrix(iRow, 11) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"            dbData.Execute \"UPDATE produtos SET ISCST = '\" & Replace(sVal, \"'\", \"''\") & \"', cClassTrib_IS = NULL, tipo_calculo_is = 1 WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n"
    b"        Else\r\n",

    "P3a: Case11 branch sem IS -> produtos"
))

# P3b - Case 11 branch CST=01:
#   apos col11="01", antes do End If -> propaga ISCST para produtos
patches.append((
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '01' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 11) = \"01\"\r\n"
    b"        End If\r\n",

    b"            dbData.Execute \"UPDATE NotaFiscalItens SET IS_CST = '01' WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 11) = \"01\"\r\n"
    b"            dbData.Execute \"UPDATE produtos SET ISCST = '01' WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n"
    b"        End If\r\n",

    "P3b: Case11 branch CST=01 -> produtos"
))

# P4a - Case 12 branch limpar (sVal=""):
#   apos cClassTrib_IS=NULL em NotaFiscalItens + cols iCol/15, antes do Else
#   -> propaga cClassTrib_IS=NULL e tipo_calculo_is=1 para produtos
patches.append((
    b"            dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, iCol) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"        Else\r\n",

    b"            dbData.Execute \"UPDATE NotaFiscalItens SET cClassTrib_IS = NULL, IS_tipo_calculo = 1, IS_vBC = 0, IS_pAliq = 0, IS_qUnid = 0, IS_vUnid = 0, IS_vIS = 0, uTrib_IS = NULL WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & Val(sItem)\r\n"
    b"            GridNotasItens.TextMatrix(iRow, iCol) = \"\"\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(0, 2)\r\n"
    b"            dbData.Execute \"UPDATE produtos SET cClassTrib_IS = NULL, tipo_calculo_is = 1 WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n"
    b"        Else\r\n",

    "P4a: Case12 branch limpar CLASS IS -> produtos"
))

# P4b - Case 12 branch valido:
#   apos cols 11/12/15, antes do End If
#   -> propaga ISCST, cClassTrib_IS e tipo_calculo_is para produtos
patches.append((
    b"            GridNotasItens.TextMatrix(iRow, 11) = sISCST2\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(curISvIS2, 2)\r\n"
    b"        End If\r\n",

    b"            GridNotasItens.TextMatrix(iRow, 11) = sISCST2\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 12) = sVal\r\n"
    b"            GridNotasItens.TextMatrix(iRow, 15) = FormatNumber(curISvIS2, 2)\r\n"
    b"            dbData.Execute \"UPDATE produtos SET ISCST = '\" & Replace(sISCST2, \"'\", \"''\") & \"', cClassTrib_IS = '\" & Replace(sVal, \"'\", \"''\") & \"', tipo_calculo_is = \" & iTipoIS2 & \" WHERE codigo = \" & Val(GridNotasItens.TextMatrix(iRow, 3))\r\n"
    b"        End If\r\n",

    "P4b: Case12 branch valido CLASS IS -> produtos"
))

errors = 0
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt} (esperado 1)")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   {desc}")

data = norm(data)

if errors:
    print(f"\n{errors} erro(s). Arquivo NAO foi salvo.")
    sys.exit(1)

with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
