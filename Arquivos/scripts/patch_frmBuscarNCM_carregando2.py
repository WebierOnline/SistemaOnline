"""
Patch parcial: apenas o trecho que falhou (CarregarGrid mostra frame)
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

patch(
    b"    ConfigurarGrid\n"
    b"    lstProdutos.rows = 1\n",
    b"    fraCarregando.Visible = True\n"
    b"    tmrCarregando.Enabled = True\n"
    b"    Me.Refresh\n"
    b"    ConfigurarGrid\n"
    b"    lstProdutos.rows = 1\n",
    "CarregarGrid mostra frame"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
