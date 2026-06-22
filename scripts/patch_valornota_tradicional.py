"""
patch_valornota_tradicional.py

Durante o periodo de transicao da reforma tributaria (2026-2032), IBS, CBS e IS
sao declarados informativamente na NF-e mas NAO compoem <vNF>.
A SEFAZ valida vNF pela formula tradicional:
  vNF = vProd + vFrete + vIPI + vSeg + vOutros + vICMSST + vFCPST - vDesc

Confirmado pelo erro na NF 246: diferenca = 71,82 = vIBS(0,17)+vCBS(1,45)+vIS(70,20).

P1: Remove + varIBS + varCBS + varIS do calculo de varNota.
P2: txtTotalDup volta a mostrar varNota (= total correto a receber do cliente,
    pois IBS/CBS ja estao embutidos no preco e IS nao e cobrado pelo varejo).
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_valornota_trad"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# P1: formula tradicional sem IBS/CBS/IS
patches.append((
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST + varIBS + varCBS + varIS - varDesconto\r\n",
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST - varDesconto\r\n"
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
