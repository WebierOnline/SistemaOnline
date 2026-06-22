"""
patch_valornota_sem_is.py
Remove varIS do calculo de varNota em AtualizarTotaisNota.

Pela especificacao NF-e com reforma tributaria, <vNF> = vProd + vFrete + vIPI
+ vSeg + vOutros + vICMSST + vIBS + vCBS - vDesc  -- sem vIS.
O IS fica em <ISTot> como bloco separado; incluir IS no varNota faz o SEFAZ
rejeitar com "Total da NF difere do somatorio dos Valores que compoe o total".

O txtTotalDup (valor a cobrar do cliente) passa a mostrar varNota + varIS
para refletir o total efetivamente devido.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_valornota_sem_is"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# P1: Remover + varIS da formula de varNota
patches.append((
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST + varIBS + varCBS + varIS - varDesconto\r\n",
    b"varNota = varProdutos + varFrete + varIPI + varSeguro + varOutras + varICMSST + varIBS + varCBS - varDesconto\r\n"
))

# P2: txtTotalDup mostra o total a cobrar = vNF + vIS (o que o cliente paga)
patches.append((
    b"txtTotalDup.Text = FormatNumber(varNota, 2)\r\n",
    b"txtTotalDup.Text = FormatNumber(varNota + varIS, 2)\r\n"
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
