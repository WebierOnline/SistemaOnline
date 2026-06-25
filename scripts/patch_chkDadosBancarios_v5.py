"""
patch_chkDadosBancarios_v5.py

Adiciona prefixo " / DADOS BANCARIO: " ao resultado de GetDadosBancariosStr.
O prefixo so e adicionado quando ha campos preenchidos (sResult nao vazio).
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\NFe_Completa.frm"
BAK = FRM + ".bak_chkDadosBancarios_v5"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

# "BANCARIO" com acento: Á = 0xC1 em cp1252
OLD = (
    b"    GetDadosBancariosStr = sResult\r\n"
    b"End Function\r\n"
)

NEW = (
    b"    If Len(sResult) > 0 Then sResult = \" / DADOS BANC\xc1RIO: \" & sResult\r\n"
    b"    GetDadosBancariosStr = sResult\r\n"
    b"End Function\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: prefixo ' / DADOS BANCARIO: ' adicionado em GetDadosBancariosStr")
print("Arquivo salvo com sucesso.")
