"""
1. Patch Produtos_Cadastro.frm: adiciona cmdBuscarNCM_Click
2. Patch os 3 VBP files: adiciona referencia a frmBuscarNCM.frm
"""
import os

# ── 1. Produtos_Cadastro.frm ─────────────────────────────────────────────────
FRM = r"C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm"
with open(FRM, "rb") as f:
    raw = f.read()
data = raw.replace(b"\r\n", b"\n").replace(b"\r", b"\n")

novo_click = (
    b"Private Sub cmdBuscarNCM_Click()\n"
    b"    With frmBuscarNCM\n"
    b"        .sCatInicial = cboCategoria.Text\n"
    b"        .sTagInicial = cboTAGs.Text\n"
    b"        .Show vbModal\n"
    b"        If .sNCMSelecionado <> \"\" Then\n"
    b"            txtNCM.Text = .sNCMSelecionado\n"
    b"            BuscarDescricaoNCM\n"
    b"        End If\n"
    b"        Unload frmBuscarNCM\n"
    b"    End With\n"
    b"End Sub\n"
    b"\n"
)

anchor = b"Private Sub cmdConsultarNCM_Click()\n"
c = data.count(anchor)
if c != 1:
    print(f"ERRO [cmdConsultarNCM anchor] ({c}x)")
else:
    data = data.replace(anchor, novo_click + anchor, 1)
    print("OK: cmdBuscarNCM_Click adicionado")

data = data.replace(b"\n", b"\r\n")
with open(FRM, "wb") as f:
    f.write(data)

# ── 2. Patch VBP files ────────────────────────────────────────────────────────
VBPS = [
    r"C:\Projeto\OnlineCommerce\OnlineCommerce.vbp",
    r"C:\Projeto\OrdemServico\OrdemServico.vbp",
    r"C:\Projeto\PDV\OnlinePDV.vbp",
]
NEW_LINE = b"Form=..\\Compartilhado\\Forms\\frmBuscarNCM.frm\r\n"
ANCHOR   = b"Form=..\\Compartilhado\\Forms\\Produtos_Cadastro.frm\r\n"

for vbp in VBPS:
    with open(vbp, "rb") as f:
        v = f.read()
    if b"frmBuscarNCM" in v:
        print(f"SKIP (ja tem): {os.path.basename(vbp)}")
        continue
    c = v.count(ANCHOR)
    if c != 1:
        print(f"ERRO [{os.path.basename(vbp)}] anchor ({c}x)")
    else:
        v = v.replace(ANCHOR, ANCHOR + NEW_LINE, 1)
        with open(vbp, "wb") as f:
            f.write(v)
        print(f"OK: {os.path.basename(vbp)}")
