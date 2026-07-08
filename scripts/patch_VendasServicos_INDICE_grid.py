"""
patch_VendasServicos_INDICE_grid.py

O grid usa sSQL diretamente. Os SQLs em cmdLocalizar_Click tem ORDER BY
hardcoded ("var_codped") ou nenhum ORDER BY. Precisamos substituir pelo
INDICE escolhido pelo usuario.

Estrategia: antes do OpenRecordset do grid, strip qualquer ORDER BY existente
no sSQL e acrescenta "ORDER BY " & Replace(INDICE, ";", "").

Para o grid o sSQL sempre usa 'pedidos AS pedidos_1', entao valores de INDICE
como 'pedidos_1.total' e 'pedidos_1.cod_pedido' sao validos diretamente.
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM, FRM + ".bak_INDICE_grid")
print(f"Backup: {FRM}.bak_INDICE_grid")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"'Debug.Print sSQL\r\n"
    b"\r\n"
    b"Set r = dbData.OpenRecordset(sSQL, totalRegistros)\r\n"
)

NEW = (
    b"'Debug.Print sSQL\r\n"
    b"\r\n"
    b"'Aplicar ordenacao pelo INDICE selecionado\r\n"
    b"Dim idxOrder As Long\r\n"
    b"idxOrder = InStrRev(sSQL, \"ORDER BY\")\r\n"
    b"If idxOrder > 0 Then sSQL = Left(sSQL, idxOrder - 1)\r\n"
    b"sSQL = sSQL & \" ORDER BY \" & Replace(INDICE, \";\", \"\")\r\n"
    b"\r\n"
    b"Set r = dbData.OpenRecordset(sSQL, totalRegistros)\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: grid sSQL agora usa INDICE no ORDER BY")
print("Arquivo salvo com sucesso.")
