"""
Fix: fraCarregando nao aparecia
1. Adiciona ZOrder 0 para garantir que o frame fique na frente do grid
2. Adiciona DoEvents apos Refresh para forcara o paint antes da query bloquear
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
    b"    fraCarregando.Visible = True\n"
    b"    tmrCarregando.Enabled = True\n"
    b"    Me.Refresh\n",
    b"    fraCarregando.Visible = True\n"
    b"    fraCarregando.ZOrder 0\n"
    b"    tmrCarregando.Enabled = True\n"
    b"    Me.Refresh\n"
    b"    DoEvents\n",
    "ZOrder + DoEvents"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
