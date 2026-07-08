"""
patch_VendasServicos_tipoVenda_totais.py

1. Remove cboTipoVenda do form definition
2. Remove PreencherTipoVendas + cboTipoVenda.ListIndex do Form_Load
3. Remove sub PreencherTipoVendas
4. Remove sub cboTipoVenda_GotFocus
5. Substituir bloco varTipoConsulta por valor fixo IN ('VENDA', 'OFICINA')
6. Adicionar lblTotalProdutos e lblTotalServicos no bloco SomaGrid
"""
import sys, shutil

FRM = r"C:\Projeto\OnlineCommerce\Forms\VendasServicos_Consulta.frm"
BAK = FRM + ".bak_tipoVenda_totais"

shutil.copy2(FRM, BAK)
print(f"Backup: {BAK}")

with open(FRM, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

patches = []

# ---------------------------------------------------------------------------
# P1: Remove cboTipoVenda do form definition
# ---------------------------------------------------------------------------
patches.append((
    b"         Begin VB.ComboBox cboTipoVenda \r\n"
    b"            Height          =   315\r\n"
    b"            Left            =   120\r\n"
    b"            TabIndex        =   53\r\n"
    b"            Top             =   420\r\n"
    b"            Width           =   2595\r\n"
    b"         End\r\n",

    b"",
    "P1: remove cboTipoVenda do form definition"
))

# ---------------------------------------------------------------------------
# P2: Remove PreencherTipoVendas + cboTipoVenda.ListIndex do Form_Load
# ---------------------------------------------------------------------------
patches.append((
    b"PreencherTipoVendas\r\n"
    b"cboTipoVenda.ListIndex = 1\r\n"
    b"\r\n",

    b"",
    "P2: remove PreencherTipoVendas e cboTipoVenda.ListIndex do Form_Load"
))

# ---------------------------------------------------------------------------
# P3: Remove sub PreencherTipoVendas
# ---------------------------------------------------------------------------
patches.append((
    b"Private Sub PreencherTipoVendas()\r\n"
    b"cboTipoVenda.Clear\r\n"
    b"cboTipoVenda.AddItem \"TODOS\"\r\n"
    b"cboTipoVenda.AddItem \"VENDAS\"\r\n"
    b"cboTipoVenda.AddItem \"OFICINA\"\r\n"
    b"End Sub\r\n",

    b"",
    "P3: remove sub PreencherTipoVendas"
))

# ---------------------------------------------------------------------------
# P4: Remove sub cboTipoVenda_GotFocus
# ---------------------------------------------------------------------------
patches.append((
    b"Private Sub cboTipoVenda_GotFocus()\r\n"
    b"moCombo.AttachTo cboTipoVenda\r\n"
    b"End Sub\r\n",

    b"",
    "P4: remove sub cboTipoVenda_GotFocus"
))

# ---------------------------------------------------------------------------
# P5: Substituir bloco varTipoConsulta por valor fixo
# ---------------------------------------------------------------------------
patches.append((
    b"'TIPO DE VENDAS\r\n"
    b"Dim varTipoConsulta As String\r\n"
    b"If cboTipoVenda.Text = \"VENDAS\" Then\r\n"
    b"   varTipoConsulta = \" = 'VENDA'\"\r\n"
    b"ElseIf cboTipoVenda.Text = \"OFICINA\" Then\r\n"
    b"   varTipoConsulta = \" = 'OFICINA'\"\r\n"
    b"ElseIf cboTipoVenda.Text = \"TODOS\" Then\r\n"
    b"   varTipoConsulta = \"IN ('VENDA', 'OFICINA')\"\r\n"
    b"End If\r\n"
    b"\r\n",

    b"Dim varTipoConsulta As String\r\n"
    b"varTipoConsulta = \"IN ('VENDA', 'OFICINA')\"\r\n"
    b"\r\n",

    "P5: varTipoConsulta fixo IN ('VENDA', 'OFICINA')"
))

# ---------------------------------------------------------------------------
# P6: Adicionar lblTotalProdutos e lblTotalServicos no bloco SomaGrid
# ---------------------------------------------------------------------------
patches.append((
    b"    lblSubtotal = Format(SomaGrid(Grid, 6), ocMONEY)\r\n"
    b"    lblTotalDesc = Format(SomaGrid(Grid, 7), ocMONEY)\r\n"
    b"    lblTotalAcresc = Format(SomaGrid(Grid, 8), ocMONEY)\r\n"
    b"    lblSubtotalBruto = Format(SomaGrid(Grid, 9), ocMONEY)\r\n",

    b"    lblTotalProdutos = Format(SomaGrid(Grid, 4), ocMONEY)\r\n"
    b"    lblTotalServicos = Format(SomaGrid(Grid, 5), ocMONEY)\r\n"
    b"    lblSubtotal = Format(SomaGrid(Grid, 6), ocMONEY)\r\n"
    b"    lblTotalDesc = Format(SomaGrid(Grid, 7), ocMONEY)\r\n"
    b"    lblTotalAcresc = Format(SomaGrid(Grid, 8), ocMONEY)\r\n"
    b"    lblSubtotalBruto = Format(SomaGrid(Grid, 9), ocMONEY)\r\n",

    "P6: adicionar lblTotalProdutos e lblTotalServicos"
))

# ---------------------------------------------------------------------------
errors = 0
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
