"""
patch_VendasServicos_criterioSec_gotfocus.py

cboCriterioSec_GotFocus: em vez de checar nomes especificos (VENDEDOR, MECANICO...),
usa cboCriterioSec.Enabled como discriminador:
  - CLIENTE           -> TODOS, MENSAL          (sub-filtro sem DATA)
  - qualquer Enabled  -> TODOS, MENSAL, DATA    (VENDEDOR, MECANICO, TECNICO, OPERADOR...)
  - Disabled          -> nao adiciona nada
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM, FRM + ".bak_criterioSec_gotfocus")
print(f"Backup: {FRM}.bak_criterioSec_gotfocus")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"cboCriterioSec.Clear\r\n"
    b"\r\n"
    b"If cboCriterioPrinc.Text = \"VENDEDOR\" Then\r\n"
    b"   cboCriterioSec.AddItem \"TODOS\"\r\n"
    b"   cboCriterioSec.AddItem \"MENSAL\"\r\n"
    b"   cboCriterioSec.AddItem \"DATA\"\r\n"
    b"ElseIf cboCriterioPrinc.Text = \"CLIENTE\" Then\r\n"
    b"   cboCriterioSec.AddItem \"TODOS\"\r\n"
    b"   cboCriterioSec.AddItem \"MENSAL\"\r\n"
    b"End If\r\n"
)

NEW = (
    b"cboCriterioSec.Clear\r\n"
    b"\r\n"
    b"If cboCriterioPrinc.Text = \"CLIENTE\" Then\r\n"
    b"   cboCriterioSec.AddItem \"TODOS\"\r\n"
    b"   cboCriterioSec.AddItem \"MENSAL\"\r\n"
    b"ElseIf cboCriterioSec.Enabled Then\r\n"
    b"   cboCriterioSec.AddItem \"TODOS\"\r\n"
    b"   cboCriterioSec.AddItem \"MENSAL\"\r\n"
    b"   cboCriterioSec.AddItem \"DATA\"\r\n"
    b"End If\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: cboCriterioSec_GotFocus agora usa cboCriterioSec.Enabled como discriminador")
print("Arquivo salvo com sucesso.")
