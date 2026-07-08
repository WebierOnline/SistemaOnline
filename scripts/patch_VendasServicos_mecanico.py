"""
patch_VendasServicos_mecanico.py

Adiciona filtro MECÂNICO ao VendasServicos_Consulta.frm:
1. PreencherPrincipal: AddItem "MECÂNICO"
2. cboCriterioPrinc_Change (topo): inicializa cboCriterioSec para MECÂNICO
3. cboCriterioPrinc_Change (ElseIf): bloco de visibilidade para MECÂNICO
   (reutiliza cboVendedor + txtCodFunc; muda Caption do lblVendedor)
   + restaura Caption "Vendedor(a):" no bloco VENDEDOR
4. cboCriterioPrinc_Change (final): adiciona MECÂNICO ao Or do cboCriterioSec
5. cmdLocalizar_Click: 3 blocos SQL (TODOS / MENSAL / DATA) via JOIN com OS
   usando OS.COD_RESPONSAVEL = txtCodFunc.Text
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
shutil.copy2(FRM, FRM + ".bak_mecanico")
print(f"Backup: {FRM}.bak_mecanico")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

# MEC\xc2NICO = MECÂNICO em cp1252
errors = 0
patches = []

# ---------------------------------------------------------------------------
# P1: PreencherPrincipal — adicionar "MECÂNICO" após MENSAL
# ---------------------------------------------------------------------------
patches.append((
    b"cboCriterioPrinc.AddItem \"MENSAL\"\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub PreencherIndice()",

    b"cboCriterioPrinc.AddItem \"MENSAL\"\r\n"
    b"cboCriterioPrinc.AddItem \"MEC\xc2NICO\"\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub PreencherIndice()",

    "P1: AddItem MECÂNICO"
))

# ---------------------------------------------------------------------------
# P2: cboCriterioPrinc_Change (topo) — inicializar cboCriterioSec para MECÂNICO
# ---------------------------------------------------------------------------
patches.append((
    b"If cboCriterioPrinc.Text = \"CLIENTE\" Then cboCriterioSec.Text = \"TODOS\"\r\n"
    b"If cboCriterioPrinc.Text = \"VENDEDOR\" Then cboCriterioSec.Text = \"TODOS\"\r\n",

    b"If cboCriterioPrinc.Text = \"CLIENTE\" Then cboCriterioSec.Text = \"TODOS\"\r\n"
    b"If cboCriterioPrinc.Text = \"VENDEDOR\" Then cboCriterioSec.Text = \"TODOS\"\r\n"
    b"If cboCriterioPrinc.Text = \"MEC\xc2NICO\" Then cboCriterioSec.Text = \"TODOS\"\r\n",

    "P2: inicializar cboCriterioSec para MECÂNICO"
))

# ---------------------------------------------------------------------------
# P3a: bloco VENDEDOR — restaurar Caption "Vendedor(a):"
# ---------------------------------------------------------------------------
patches.append((
    b"   If cboCriterioPrinc.Text = \"VENDEDOR\" And cboCriterioSec.Text = \"TODOS\" Then\r\n"
    b"      lblVendedor.Visible = True\r\n"
    b"      cboVendedor.Visible = True\r\n",

    b"   If cboCriterioPrinc.Text = \"VENDEDOR\" And cboCriterioSec.Text = \"TODOS\" Then\r\n"
    b"      lblVendedor.Caption = \"Vendedor(a):\"\r\n"
    b"      lblVendedor.Visible = True\r\n"
    b"      cboVendedor.Visible = True\r\n",

    "P3a: restaurar Caption Vendedor(a): no bloco VENDEDOR"
))

# ---------------------------------------------------------------------------
# P3b: inserir bloco ElseIf MECÂNICO antes do Else/Exit Sub
# O bloco MENSAL termina com: If cboMes.Visible = True Then cboMes.SetFocus
# ---------------------------------------------------------------------------
patches.append((
    b"      If cboMes.Visible = True Then cboMes.SetFocus\r\n"
    b"   Else\r\n"
    b"      Exit Sub\r\n"
    b"   End If\r\n"
    b"   \r\n"
    b"   \'cboCriterioSec.Clear\r\n",

    b"      If cboMes.Visible = True Then cboMes.SetFocus\r\n"
    b"   ElseIf cboCriterioPrinc.Text = \"MEC\xc2NICO\" And cboCriterioSec.Text = \"TODOS\" Then\r\n"
    b"      lblVendedor.Caption = \"Mec\xc2nico:\"\r\n"
    b"      lblVendedor.Visible = True\r\n"
    b"      cboVendedor.Visible = True\r\n"
    b"      \r\n"
    b"      lblInicio.Visible = False\r\n"
    b"      mskInicio.Visible = False\r\n"
    b"      lblFim.Visible = False\r\n"
    b"      mskFim.Visible = False\r\n"
    b"      lblAte.Visible = False\r\n"
    b"      cmdCalendario1.Visible = False\r\n"
    b"      cmdCalendario2.Visible = False\r\n"
    b"    cmdCal1.Visible = False\r\n"
    b"    mskData.Visible = False\r\n"
    b"    lblData.Visible = False\r\n"
    b"      \r\n"
    b"      lblClientes.Visible = False\r\n"
    b"      cboCliente.Visible = False\r\n"
    b"      \r\n"
    b"      lblCodigo.Visible = False\r\n"
    b"      txtCodigo.Visible = False\r\n"
    b"      \r\n"
    b"      lblMes.Visible = False\r\n"
    b"      cboMes.Visible = False\r\n"
    b"      lblAno.Visible = False\r\n"
    b"      cboAno.Visible = False\r\n"
    b"      \r\n"
    b"      cboCriterioSec.Enabled = True\r\n"
    b"      lblSubConsulta.Enabled = True\r\n"
    b"    lblVendedor.Top = 180\r\n"
    b"    cboVendedor.Top = 420\r\n"
    b"    cmdLocalizar.Top = 420\r\n"
    b"    cmdLocalizar.Left = 5220\r\n"
    b"      \r\n"
    b"      cboVendedor.SetFocus\r\n"
    b"   Else\r\n"
    b"      Exit Sub\r\n"
    b"   End If\r\n"
    b"   \r\n"
    b"   \'cboCriterioSec.Clear\r\n",

    "P3b: ElseIf MECÂNICO em cboCriterioPrinc_Change"
))

# ---------------------------------------------------------------------------
# P4: condição final cboCriterioPrinc_Change — adicionar MECÂNICO ao Or
# ---------------------------------------------------------------------------
patches.append((
    b"   If cboCriterioPrinc.Text = \"VENDEDOR\" Or cboCriterioPrinc.Text = \"CLIENTE\" Then\r\n"
    b"      cboCriterioSec.Text = \"TODOS\"\r\n"
    b"   Else\r\n"
    b"      cboCriterioSec.Text = \"\"\r\n"
    b"   End If\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub cboCriterioPrinc_Click()",

    b"   If cboCriterioPrinc.Text = \"VENDEDOR\" Or cboCriterioPrinc.Text = \"CLIENTE\" Or cboCriterioPrinc.Text = \"MEC\xc2NICO\" Then\r\n"
    b"      cboCriterioSec.Text = \"TODOS\"\r\n"
    b"   Else\r\n"
    b"      cboCriterioSec.Text = \"\"\r\n"
    b"   End If\r\n"
    b"End Sub\r\n"
    b"\r\n"
    b"Private Sub cboCriterioPrinc_Click()",

    "P4: adicionar MECÂNICO ao Or final de cboCriterioPrinc_Change"
))

# ---------------------------------------------------------------------------
# P5: cmdLocalizar_Click — 3 blocos SQL para MECÂNICO antes de 'CLIENTE - TODOS
# SQL: JOIN com OS usando OS.COD_RESPONSAVEL = txtCodFunc.Text
# ---------------------------------------------------------------------------
patches.append((
    b"            \"ORDER BY var_codped\"\r\n"
    b"\r\n"
    b"   \'CLIENTE - TODOS\r\n",

    b"            \"ORDER BY var_codped\"\r\n"
    b"\r\n"
    b"   'MEC\xc2NICO - TODOS\r\n"
    b"    ElseIf cboCriterioPrinc.Text = \"MEC\xc2NICO\" And cboCriterioSec.Text = \"TODOS\" Then\r\n"
    b"        If cboVendedor.Text = \"\" Then Limpar_Grid_Venda: Exit Sub\r\n"
    b"\r\n"
    b"        sSQL = \"SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE,  \" & _\r\n"
    b"            \"(SELECT DISTINCT cliente.Nome FROM cliente WHERE cliente.CODIGO = pedidos_1.COD_CLIENTE) AS Nome \" & _\r\n"
    b"            \"FROM pedidos AS pedidos_1 INNER JOIN OS ON OS.COD_PEDIDO = pedidos_1.COD_PEDIDO \" & _\r\n"
    b"            \"WHERE (OS.COD_RESPONSAVEL = \" & txtCodFunc.Text & \") AND (pedidos_1.tipo_pedido \" & varTipoConsulta & \") AND (pedidos_1.CANCELADO = 0) \" & TipoPgto & \" \" & _\r\n"
    b"            \"ORDER BY var_codped\"\r\n"
    b"\r\n"
    b"   'MEC\xc2NICO - MENSAL\r\n"
    b"    ElseIf cboCriterioPrinc.Text = \"MEC\xc2NICO\" And cboCriterioSec.Text = \"MENSAL\" Then\r\n"
    b"        If cboVendedor.Text = \"\" Then Limpar_Grid_Venda: Exit Sub\r\n"
    b"        If cboMes.Text = \"\" Or cboAno.Text = \"\" Then Limpar_Grid_Venda: Exit Sub\r\n"
    b"\r\n"
    b"        sSQL = \"SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE,  \" & _\r\n"
    b"            \"(SELECT DISTINCT cliente.Nome FROM cliente WHERE cliente.CODIGO = pedidos_1.COD_CLIENTE) AS Nome \" & _\r\n"
    b"            \"FROM pedidos AS pedidos_1 INNER JOIN OS ON OS.COD_PEDIDO = pedidos_1.COD_PEDIDO \" & _\r\n"
    b"            \"WHERE (OS.COD_RESPONSAVEL = \" & txtCodFunc.Text & \") AND (Month(pedidos_1.data_compra) = \" & cboMes.ListIndex + 1 & \") And (Year(pedidos_1.data_compra) = \" & cboAno & \") AND (pedidos_1.tipo_pedido \" & varTipoConsulta & \") AND (pedidos_1.CANCELADO = 0) \" & TipoPgto & \" \" & _\r\n"
    b"            \"ORDER BY var_codped\"\r\n"
    b"\r\n"
    b"   'MEC\xc2NICO - DATA\r\n"
    b"    ElseIf cboCriterioPrinc.Text = \"MEC\xc2NICO\" And cboCriterioSec.Text = \"DATA\" Then\r\n"
    b"        If cboVendedor.Text = \"\" Then Limpar_Grid_Venda: Exit Sub\r\n"
    b"        If mskData.Text = \"\" Then Exit Sub\r\n"
    b"\r\n"
    b"        sSQL = \"SELECT DISTINCT pedidos_1.COD_PEDIDO AS var_codped, pedidos_1.DATA_COMPRA, pedidos_1.SUBTOTAL, pedidos_1.ValorAcrescReal, pedidos_1.ValorDescReal, pedidos_1.TOTAL AS var_total, pedidos_1.TIPO_PAGAMENTO, pedidos_1.PAGAMENTO, pedidos_1.TIPO_PEDIDO, pedidos_1.COD_CLIENTE,  \" & _\r\n"
    b"            \"(SELECT DISTINCT cliente.Nome FROM cliente WHERE cliente.CODIGO = pedidos_1.COD_CLIENTE) AS Nome \" & _\r\n"
    b"            \"FROM pedidos AS pedidos_1 INNER JOIN OS ON OS.COD_PEDIDO = pedidos_1.COD_PEDIDO \" & _\r\n"
    b"            \"WHERE (OS.COD_RESPONSAVEL = \" & txtCodFunc.Text & \") AND (pedidos_1.data_compra = CONVERT(DATETIME, '\" & Format(mskData, ocDATA) & \"', 103)) AND (pedidos_1.tipo_pedido \" & varTipoConsulta & \") AND (pedidos_1.CANCELADO = 0) \" & TipoPgto & \" \" & _\r\n"
    b"            \"ORDER BY var_codped\"\r\n"
    b"\r\n"
    b"   \'CLIENTE - TODOS\r\n",

    "P5: blocos SQL MECÂNICO (TODOS/MENSAL/DATA) em cmdLocalizar_Click"
))

# ---------------------------------------------------------------------------
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt} (esperado 1)")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   {desc}")

if errors:
    print(f"\n{errors} erro(s). Arquivo NAO salvo.")
    sys.exit(1)

data = norm(data)
with open(FRM, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
