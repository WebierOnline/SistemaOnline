"""
patch_VendasServicos_imprimir_mecanico.py

cmdImprimir_Click: adiciona ElseIf generico antes do Else final do bloco rfCons1.
Usa cboCriterioSec.Enabled como discriminador (mesma logica do GotFocus):
  - cobre MECANICO, TECNICO, OPERADOR e qualquer nome futuro que habilite o combo.
  - exibe: "<nome do criterio> = <funcionario selecionado>"
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM, FRM + ".bak_imprimir_mecanico")
print(f"Backup: {FRM}.bak_imprimir_mecanico")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

OLD = (
    b"   ElseIf cboCriterioPrinc.Text = \"TODOS\" Then\r\n"
    b"      REL_Cons_Venda.rfCons1.Caption = \"TODOS\"\r\n"
    b"   Else\r\n"
    b"      REL_Cons_Venda.rfCons1.Caption = \"TODOS\"\r\n"
    b"   End If\r\n"
)

NEW = (
    b"   ElseIf cboCriterioPrinc.Text = \"TODOS\" Then\r\n"
    b"      REL_Cons_Venda.rfCons1.Caption = \"TODOS\"\r\n"
    b"   ElseIf cboCriterioSec.Enabled Then\r\n"
    b"      REL_Cons_Venda.rfCons1.Caption = cboCriterioPrinc.Text & \" = \" & cboVendedor.Text\r\n"
    b"   Else\r\n"
    b"      REL_Cons_Venda.rfCons1.Caption = \"TODOS\"\r\n"
    b"   End If\r\n"
)

cnt = data.count(OLD)
if cnt != 1:
    print(f"ERRO: trecho encontrado {cnt}x (esperado 1). Arquivo NAO alterado.")
    sys.exit(1)

data = data.replace(OLD, NEW)
data = norm(data)

with open(FRM, "wb") as f:
    f.write(data)

print("OK: rfCons1 agora exibe criterio generico para MECANICO/TECNICO/OPERADOR")
print("Arquivo salvo com sucesso.")
