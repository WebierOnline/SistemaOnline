"""
patch_p1_calcular_is_qtrib.py

Fix minimo: remove zeragem de curISqUnid para tipo=1 (ad valorem) em Calcular_IS.
O schema NF-e exige qTrib > 0 SEMPRE que o bloco IS esta presente.
Apenas curISvUnid (preco unitario) nao se aplica para ad valorem / sem IS.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_p1_calcular_is"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"    ' Ad Valorem e sem IS: qUnid e vUnid nao se aplicam (zerar para evitar rejeicao SEFAZ)\r\n"
    b"    If iTipoIS = 0 Or iTipoIS = 1 Then\r\n"
    b"        curISqUnid = 0\r\n"
    b"        curISvUnid = 0\r\n"
    b"    End If\r\n"
)

NEW = (
    b"    ' Schema exige qTrib > 0 sempre; apenas vUnid nao se aplica em ad valorem / sem IS\r\n"
    b"    If iTipoIS = 0 Or iTipoIS = 1 Then\r\n"
    b"        curISvUnid = 0\r\n"
    b"    End If\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: Calcular_IS - curISqUnid nao mais zerado para tipo=1")
print("Arquivo salvo com sucesso.")
