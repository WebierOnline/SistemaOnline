"""
patch_revert_valornota_e_cbs_vbc.py

P1: Reverte remocao de varIS do varNota — o <vNF> da NF-e COM reforma tributaria
    DEVE incluir IS. A rejeicao original era CBS subtotalizado (vCBS=1,41 em vez
    de 1,45 devido a base duplo-descontada), nao a presenca do IS.
    Formula correta: vNF = vProd + vFrete + vIPI + vSeg + vOutros + vICMSST
                          + vIBS + vCBS + vIS - vDesc

P2: Reverte txtTotalDup — volta a exibir so varNota (que ja inclui IS).

P3: Adiciona CBS_vBC ao UPDATE de RecalcularItensNota, usando o mesmo
    curBCCBSIBS2 do IBS. Corrige o valor armazenado (era double-discounted:
    IBS_vBC - ValorDesconto). Tambem evita re-ocorrencia em novas notas.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_revert_valornota"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# P1: Restaurar + varIS na formula de varNota
patches.append((
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST + varIBS + varCBS - varDesconto\r\n",
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST + varIBS + varCBS + varIS - varDesconto\r\n"
))

# P2: Reverter txtTotalDup — varNota ja inclui IS, nao precisa somar de novo
patches.append((
    b"txtTotalDup.Text = FormatNumber(varNota + varIS, 2)\r\n",
    b"txtTotalDup.Text = FormatNumber(varNota, 2)\r\n"
))

# P3: Adicionar CBS_vBC = curBCCBSIBS2 no UPDATE de RecalcularItensNota
patches.append((
    b"           \"IBS_vBC = \" & FSQL(curBCCBSIBS2, 2) & \", \" & _\r\n",
    b"           \"IBS_vBC = \" & FSQL(curBCCBSIBS2, 2) & \", \" & _\r\n"
    b"           \"CBS_vBC = \" & FSQL(curBCCBSIBS2, 2) & \", \" & _\r\n"
))

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
    print(f"\n{errors} erro(s). Arquivo NAO foi salvo.")
    sys.exit(1)

with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
