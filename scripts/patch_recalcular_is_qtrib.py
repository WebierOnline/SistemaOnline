"""
patch_recalcular_is_qtrib.py

Corrige dois bugs relacionados a qTrib/uTrib no IS da reforma tributaria:

1. Calcular_IS: nao mais zera curISqUnid para tipo=1 (ad valorem).
   O schema exige qTrib > 0 SEMPRE que o bloco IS esta presente.
   Apenas curISvUnid (preco unit.) e zerado para tipos 0 e 1.

2. RecalcularItensNota: passa a corrigir IS_qUnid (quando zerado) usando
   QuantidadeComercial, e busca uTrib_IS de tbISClassTrib, salvando ambos
   no UPDATE. Permite recuperar NFs com dados IS invalidos via botao Recalcular.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_recalcular_is_qtrib"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ---------------------------------------------------------------------------
# P1 - Calcular_IS: remover zeragem de curISqUnid para tipo=1
# ---------------------------------------------------------------------------
patches.append((
    b"    ' Ad Valorem e sem IS: qUnid e vUnid nao se aplicam (zerar para evitar rejeicao SEFAZ)\r\n"
    b"    If iTipoIS = 0 Or iTipoIS = 1 Then\r\n"
    b"        curISqUnid = 0\r\n"
    b"        curISvUnid = 0\r\n"
    b"    End If\r\n",

    b"    ' Schema exige qTrib > 0 sempre; apenas vUnid nao se aplica em ad valorem / sem IS\r\n"
    b"    If iTipoIS = 0 Or iTipoIS = 1 Then\r\n"
    b"        curISvUnid = 0\r\n"
    b"    End If\r\n",

    "P1: Calcular_IS - manter qUnid para tipo=1"
))

# ---------------------------------------------------------------------------
# P2 - RecalcularItensNota: adicionar QuantidadeComercial e cClassTrib_IS ao SELECT
# ---------------------------------------------------------------------------
patches.append((
    b"       \"IS_tipo_calculo, IS_qUnid, IS_vUnid \" & _\r\n"
    b"       \"FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)\r\n",

    b"       \"IS_tipo_calculo, IS_qUnid, IS_vUnid, QuantidadeComercial, cClassTrib_IS \" & _\r\n"
    b"       \"FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)\r\n",

    "P2: RecalcularItensNota SELECT + QuantidadeComercial/cClassTrib_IS"
))

# ---------------------------------------------------------------------------
# P3 - Declarar variaveis novas no cabecalho de RecalcularItensNota
# ---------------------------------------------------------------------------
patches.append((
    b"Dim curISqUnid2  As Currency, curISvUnid2 As Currency\r\n",

    b"Dim curISqUnid2  As Currency, curISvUnid2 As Currency\r\n"
    b"Dim dblQtdeIS2 As Double, sClassTribIS2 As String, sUTribIS2 As String, sSQLUTrib2 As String\r\n",

    "P3: RecalcularItensNota Dim novas variaveis"
))

# ---------------------------------------------------------------------------
# P4 - Loop: apos ler curISqUnid2/curISvUnid2, corrigir qUnid e buscar uTrib_IS
# Obs: SQLExecutaRetorno numa unica linha para evitar quebra sem continuacao VB6
# ---------------------------------------------------------------------------
patches.append((
    b"    curISqUnid2 = CCur(IIf(IsNull(rItens(\"IS_qUnid\")), 0, rItens(\"IS_qUnid\")))\r\n"
    b"    curISvUnid2 = CCur(IIf(IsNull(rItens(\"IS_vUnid\")), 0, rItens(\"IS_vUnid\")))\r\n"
    b"    Select Case iTipoIS2\r\n",

    b"    curISqUnid2 = CCur(IIf(IsNull(rItens(\"IS_qUnid\")), 0, rItens(\"IS_qUnid\")))\r\n"
    b"    curISvUnid2 = CCur(IIf(IsNull(rItens(\"IS_vUnid\")), 0, rItens(\"IS_vUnid\")))\r\n"
    b"    dblQtdeIS2 = CDbl(IIf(IsNull(rItens(\"QuantidadeComercial\")), 0, rItens(\"QuantidadeComercial\")))\r\n"
    b"    sClassTribIS2 = Trim(IIf(IsNull(rItens(\"cClassTrib_IS\")), \"\", rItens(\"cClassTrib_IS\")))\r\n"
    b"    ' qUnid obrigatorio pelo schema (qTrib > 0); corrigir quando foi zerado incorretamente\r\n"
    b"    If curISqUnid2 = 0 And dblQtdeIS2 > 0 Then curISqUnid2 = CCur(Format(dblQtdeIS2, \"0.0000\"))\r\n"
    b"    sUTribIS2 = \"\"\r\n"
    b"    If sClassTribIS2 <> \"\" Then\r\n"
    b"        sUTribIS2 = SQLExecutaRetorno(\"SELECT TOP 1 uTrib_IS FROM tbISClassTrib WHERE cClassTrib_IS = '\" & Replace(sClassTribIS2, \"'\", \"''\") & \"' ORDER BY CASE WHEN GETDATE() BETWEEN dIniVig AND dFimVig THEN 0 ELSE 1 END ASC, dFimVig DESC\", \"uTrib_IS\", \"\")\r\n"
    b"    End If\r\n"
    b"    If sUTribIS2 = \"\" Then sSQLUTrib2 = \"NULL\" Else sSQLUTrib2 = \"'\" & Replace(sUTribIS2, \"'\", \"''\") & \"'\"\r\n"
    b"    Select Case iTipoIS2\r\n",

    "P4: RecalcularItensNota loop - corrigir qUnid e buscar uTrib_IS"
))

# ---------------------------------------------------------------------------
# P5 - RecalcularItensNota UPDATE: adicionar IS_qUnid e uTrib_IS
# ---------------------------------------------------------------------------
patches.append((
    b"           \"IS_vBC = \" & FSQL(curBCIS2, 2) & \", \" & _\r\n"
    b"           \"IS_vIS = \" & FSQL(curVIS2, 2) & \" \" & _\r\n"
    b"           \"WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & vItem\r\n",

    b"           \"IS_vBC = \" & FSQL(curBCIS2, 2) & \", \" & _\r\n"
    b"           \"IS_qUnid = \" & FSQL(CDbl(curISqUnid2), 4) & \", \" & _\r\n"
    b"           \"IS_vIS = \" & FSQL(curVIS2, 2) & \", \" & _\r\n"
    b"           \"uTrib_IS = \" & sSQLUTrib2 & \" \" & _\r\n"
    b"           \"WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & vItem\r\n",

    "P5: RecalcularItensNota UPDATE + IS_qUnid + uTrib_IS"
))

# ---------------------------------------------------------------------------
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
