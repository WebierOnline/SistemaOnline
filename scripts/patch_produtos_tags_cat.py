"""
Patch cboTAGs_GotFocus: filtra tags pela categoria selecionada
"""
FILE = r"C:\Projeto\Compartilhado\Forms\Produtos_Cadastro.frm"
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
    b"RsOpen rTags, \"SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' ORDER BY TAGS\"\n",
    b"RsOpen rTags, \"SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' AND categoria = '\" & Replace(cboCategoria.Text, \"'\", \"''\") & \"' ORDER BY TAGS\"\n",
    "filtro por categoria"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
