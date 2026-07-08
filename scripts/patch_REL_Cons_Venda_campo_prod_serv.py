"""
patch_REL_Cons_Venda_campo_prod_serv.py

1. VendasServicos_Consulta.frm
   - Substitui 'printSQL = sSQL' por um wrapper que acrescenta vProd e vServ
     ao SELECT via subquery correlated, sem tocar nas +10 variantes de SQL.
   - Remove o ORDER BY do sSQL original (InStrRev) e reaplica na query externa.

2. REL_Cons_Venda.frm
   - Adiciona Campo = "vProd" no ReportField7
   - Adiciona Campo = "vServ" no ReportField9
"""
import sys, shutil

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

errors = 0

# ===========================================================================
# 1. VendasServicos_Consulta.frm — wrapper printSQL
# ===========================================================================
FRM1 = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM1, FRM1 + ".bak_printSQL_wrap")
print(f"Backup: {FRM1}.bak_printSQL_wrap")

with open(FRM1, "rb") as f:
    d1 = f.read()

OLD1 = b"printSQL = sSQL\r\nEnd Sub\r\n"

NEW1 = (
    b"'Envolver sSQL num subquery para adicionar vProd e vServ sem alterar cada variante de SQL\r\n"
    b"Dim sPrintBase As String\r\n"
    b"Dim idxOrd As Long\r\n"
    b"sPrintBase = sSQL\r\n"
    b"idxOrd = InStrRev(sPrintBase, \"ORDER BY\")\r\n"
    b"If idxOrd > 0 Then sPrintBase = Left(sPrintBase, idxOrd - 1)\r\n"
    b"printSQL = \"SELECT *, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(pi.total),0) FROM pedidos_itens pi WHERE pi.cod_pedido=q.var_codped) AS vProd, \" & _\r\n"
    b"    \"(SELECT ISNULL(SUM(sv.total),0) FROM OS_Servicos_Auto sv INNER JOIN OS ON sv.cod_os=OS.COD_OS WHERE OS.COD_PEDIDO=q.var_codped) AS vServ \" & _\r\n"
    b"    \"FROM (\" & sPrintBase & \") AS q ORDER BY q.var_codped\"\r\n"
    b"End Sub\r\n"
)

cnt = d1.count(OLD1)
if cnt != 1:
    print(f"ERRO VendasServicos printSQL: count={cnt} (esperado 1)")
    errors += 1
else:
    d1 = d1.replace(OLD1, NEW1)
    print("OK   VendasServicos: printSQL envolto com vProd/vServ")
    d1 = norm(d1)
    with open(FRM1, "wb") as f:
        f.write(d1)
    print("     Arquivo salvo.")

# ===========================================================================
# 2. REL_Cons_Venda.frm — Campo nos ReportField7 e ReportField9
# ===========================================================================
FRM2 = r"C:\Projeto\OnlineCommerce\Forms\REL_Cons_Venda.frm"
shutil.copy2(FRM2, FRM2 + ".bak_campo_prod_serv")
print(f"Backup: {FRM2}.bak_campo_prod_serv")

with open(FRM2, "rb") as f:
    d2 = f.read()

# --- ReportField7: Campo = "vProd" ---
OLD2a = (
    b"         _ExtentX        =   1931\r\n"
    b"         _ExtentY        =   344\r\n"
    b"         Formato         =   \"##,##0.00\"\r\n"
    b"         TipoCampo       =   1\r\n"
    b"         Alignment       =   1\r\n"
)
NEW2a = (
    b"         _ExtentX        =   1931\r\n"
    b"         _ExtentY        =   344\r\n"
    b"         Campo           =   \"vProd\"\r\n"
    b"         Formato         =   \"##,##0.00\"\r\n"
    b"         TipoCampo       =   1\r\n"
    b"         Alignment       =   1\r\n"
)

cnt = d2.count(OLD2a)
if cnt != 1:
    print(f"ERRO REL RF7 Campo: count={cnt} (esperado 1)")
    errors += 1
else:
    d2 = d2.replace(OLD2a, NEW2a)
    print("OK   REL_Cons_Venda: ReportField7 Campo = vProd")

# --- ReportField9: Campo = "vServ" ---
OLD2b = (
    b"         _ExtentX        =   1720\r\n"
    b"         _ExtentY        =   344\r\n"
    b"         Formato         =   \"##,##0.00\"\r\n"
    b"         TipoCampo       =   1\r\n"
    b"         Alignment       =   1\r\n"
)
NEW2b = (
    b"         _ExtentX        =   1720\r\n"
    b"         _ExtentY        =   344\r\n"
    b"         Campo           =   \"vServ\"\r\n"
    b"         Formato         =   \"##,##0.00\"\r\n"
    b"         TipoCampo       =   1\r\n"
    b"         Alignment       =   1\r\n"
)

cnt = d2.count(OLD2b)
if cnt != 1:
    print(f"ERRO REL RF9 Campo: count={cnt} (esperado 1)")
    errors += 1
else:
    d2 = d2.replace(OLD2b, NEW2b)
    print("OK   REL_Cons_Venda: ReportField9 Campo = vServ")

if not errors:
    d2 = norm(d2)
    with open(FRM2, "wb") as f:
        f.write(d2)
    print("     Arquivo salvo.")

if errors:
    print(f"\n{errors} erro(s) encontrado(s).")
    sys.exit(1)
else:
    print("\nTudo aplicado com sucesso.")
