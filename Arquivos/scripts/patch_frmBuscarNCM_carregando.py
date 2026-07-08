"""
Patch frmBuscarNCM.frm: label animado "Carregando..."
1. Adiciona fraCarregando (Frame com lblCarregando) e tmrCarregando na definicao do form
2. Adiciona Dim iDots As Integer
3. Form_Load: inicializa fraCarregando.Visible=False e iDots=0
4. CarregarGrid: mostra/esconde frame e timer em volta da carga
5. Adiciona tmrCarregando_Timer (anima os pontos)
"""
FILE = r"C:\Projeto\Compartilhado\Forms\frmBuscarNCM.frm"
with open(FILE, "rb") as f:
    raw = f.read()
data = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")

errors = 0

def patch(old, new, label):
    global data, errors
    c = data.count(old)
    if c != 1:
        print(f"ERRO [{label}] ({c}x)")
        errors += 1
    else:
        data = data.replace(old, new, 1)
        print(f"OK: {label}")

# 1. Adiciona fraCarregando + tmrCarregando na definicao do form
patch(
    b"End\n"
    b"Attribute VB_Name = \"frmBuscarNCM\"\n",
    b"   Begin VB.Frame fraCarregando \n"
    b"      BackColor       =   &H00FFFFFF&\n"
    b"      BorderStyle     =   1  'Fixed Single\n"
    b"      Caption         =   \"\"\n"
    b"      Height          =   600\n"
    b"      Left            =   3090\n"
    b"      TabIndex        =   50\n"
    b"      Top             =   3200\n"
    b"      Visible         =   0   'False\n"
    b"      Width           =   4800\n"
    b"      Begin VB.Label lblCarregando \n"
    b"         Alignment       =   2  'Center\n"
    b"         BackStyle       =   0  'Transparent\n"
    b"         Caption         =   \"Carregando...\"\n"
    b"         BeginProperty Font \n"
    b"            Name            =   \"MS Sans Serif\"\n"
    b"            Size            =   14\n"
    b"            Charset         =   0\n"
    b"            Weight          =   700\n"
    b"            Underline       =   0   'False\n"
    b"            Italic          =   0   'False\n"
    b"            Strikethrough   =   0   'False\n"
    b"         EndProperty\n"
    b"         Height          =   420\n"
    b"         Left            =   60\n"
    b"         TabIndex        =   0\n"
    b"         Top             =   90\n"
    b"         Width           =   4680\n"
    b"      End\n"
    b"   End\n"
    b"   Begin VB.Timer tmrCarregando \n"
    b"      Enabled         =   0   'False\n"
    b"      Interval        =   400\n"
    b"      Left            =   120\n"
    b"      Top             =   6840\n"
    b"   End\n"
    b"End\n"
    b"Attribute VB_Name = \"frmBuscarNCM\"\n",
    "fraCarregando + tmrCarregando definicao"
)

# 2. Variavel iDots apos sTagInicial
patch(
    b"Public sTagInicial As String\n",
    b"Public sTagInicial As String\n"
    b"Private iDots As Integer\n",
    "iDots var"
)

# 3. Form_Load: inicializa frame invisivel
patch(
    b"    sNCMSelecionado = \"\"\n"
    b"    'optDesc.Value = True\n",
    b"    sNCMSelecionado = \"\"\n"
    b"    iDots = 0\n"
    b"    fraCarregando.Visible = False\n"
    b"    'optDesc.Value = True\n",
    "Form_Load init"
)

# 4a. CarregarGrid: mostra frame antes de carregar
patch(
    b"    ConfigurarGrid\n"
    b"    lstProdutos.Rows = 1\n",
    b"    fraCarregando.Visible = True\n"
    b"    tmrCarregando.Enabled = True\n"
    b"    Me.Refresh\n"
    b"    ConfigurarGrid\n"
    b"    lstProdutos.Rows = 1\n",
    "CarregarGrid mostra frame"
)

# 4b. CarregarGrid: esconde frame apos carregar
patch(
    b"    If rProd.State <> 0 Then rProd.Close\n"
    b"End Sub\n"
    b"\n"
    b"' --- Eventos dos combos ---\n",
    b"    If rProd.State <> 0 Then rProd.Close\n"
    b"    tmrCarregando.Enabled = False\n"
    b"    fraCarregando.Visible = False\n"
    b"End Sub\n"
    b"\n"
    b"' --- Eventos dos combos ---\n",
    "CarregarGrid esconde frame"
)

# 5. Adiciona tmrCarregando_Timer ao final
patch(
    b"Private Sub cmdFechar_Click()\n"
    b"    sNCMSelecionado = \"\"\n"
    b"    Me.Hide\n"
    b"End Sub\n",
    b"Private Sub cmdFechar_Click()\n"
    b"    sNCMSelecionado = \"\"\n"
    b"    Me.Hide\n"
    b"End Sub\n"
    b"\n"
    b"Private Sub tmrCarregando_Timer()\n"
    b"    iDots = (iDots Mod 3) + 1\n"
    b"    lblCarregando.Caption = \"Carregando\" & String(iDots, \".\")\n"
    b"End Sub\n",
    "tmrCarregando_Timer"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
