"""
patch_recalc_is_qunid_simples.py

Versao simplificada: apenas corrige IS_qUnid em RecalcularItensNota.

P2: Adiciona QuantidadeComercial ao SELECT (necessario para corrigir qUnid=0)
P4: Apos ler curISqUnid2, se for 0 usa QuantidadeComercial diretamente do recordset
P5: Adiciona IS_qUnid ao UPDATE (para persistir a correcao)

Nao adiciona novas Dim, nao faz lookup em tbISClassTrib.
uTrib_IS nao e tocado pelo Recalcular (sobrevive ao UPDATE).
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_recalc_qunid_simples"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# P2: Add QuantidadeComercial to SELECT
patches.append((
    b"       \"IS_tipo_calculo, IS_qUnid, IS_vUnid \" & _\r\n"
    b"       \"FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)\r\n",

    b"       \"IS_tipo_calculo, IS_qUnid, IS_vUnid, QuantidadeComercial \" & _\r\n"
    b"       \"FROM NotaFiscalItens WHERE CodigoNota = \" & Val(txtCodNota.Text)\r\n",

    "P2: SELECT + QuantidadeComercial"
))

# P4: apos ler curISqUnid2/vUnid2, corrigir se for 0
patches.append((
    b"    curISqUnid2 = CCur(IIf(IsNull(rItens(\"IS_qUnid\")), 0, rItens(\"IS_qUnid\")))\r\n"
    b"    curISvUnid2 = CCur(IIf(IsNull(rItens(\"IS_vUnid\")), 0, rItens(\"IS_vUnid\")))\r\n"
    b"    Select Case iTipoIS2\r\n",

    b"    curISqUnid2 = CCur(IIf(IsNull(rItens(\"IS_qUnid\")), 0, rItens(\"IS_qUnid\")))\r\n"
    b"    curISvUnid2 = CCur(IIf(IsNull(rItens(\"IS_vUnid\")), 0, rItens(\"IS_vUnid\")))\r\n"
    b"    If curISqUnid2 = 0 Then\r\n"
    b"        If Not IsNull(rItens(\"QuantidadeComercial\")) Then\r\n"
    b"            If CDbl(rItens(\"QuantidadeComercial\")) > 0 Then curISqUnid2 = CCur(rItens(\"QuantidadeComercial\"))\r\n"
    b"        End If\r\n"
    b"    End If\r\n"
    b"    Select Case iTipoIS2\r\n",

    "P4: corrigir curISqUnid2=0 com QuantidadeComercial"
))

# P5: adicionar IS_qUnid ao UPDATE
patches.append((
    b"           \"IS_vBC = \" & FSQL(curBCIS2, 2) & \", \" & _\r\n"
    b"           \"IS_vIS = \" & FSQL(curVIS2, 2) & \" \" & _\r\n"
    b"           \"WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & vItem\r\n",

    b"           \"IS_vBC = \" & FSQL(curBCIS2, 2) & \", \" & _\r\n"
    b"           \"IS_qUnid = \" & FSQL(CDbl(curISqUnid2), 4) & \", \" & _\r\n"
    b"           \"IS_vIS = \" & FSQL(curVIS2, 2) & \" \" & _\r\n"
    b"           \"WHERE CodigoNota = \" & Val(txtCodNota.Text) & \" AND ITEM = \" & vItem\r\n",

    "P5: UPDATE + IS_qUnid"
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
