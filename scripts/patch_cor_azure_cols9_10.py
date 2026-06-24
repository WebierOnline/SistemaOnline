"""
patch_cor_azure_cols9_10.py
Move cols 9 e 10 do array amarelo (editável) para o array azure (reforma),
mantendo o comportamento de edição intacto — só muda a cor de fundo.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_cor_azure"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = [
    # Tira 9 e 10 do array amarelo
    (
        b"      For Each colEdit In Array(2, 5, 6, 7, 8, 9, 10, 22, 24, 26, 28, 29, 30)",
        b"      For Each colEdit In Array(2, 5, 6, 7, 8, 22, 24, 26, 28, 29, 30)"
    ),
    # Adiciona 9 e 10 ao array azure (reforma)
    (
        b"      For Each colRef In Array(11, 12, 13)",
        b"      For Each colRef In Array(9, 10, 11, 12, 13)"
    ),
]

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
