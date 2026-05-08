"""
Fixes em frmBuscarNCM.frm:
1. CarregarCategorias: SELECT DISTINCT (elimina repeticoes)
2. CarregarTags: primeiro item "TODOS" (em vez de vazio)
3. CarregarGrid: campos de texto desvinculam combos; busca por descricao = LIKE %palavra%
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

# 1. DISTINCT nas categorias
patch(
    b"    RsOpen rCat, \"SELECT Categoria FROM Categorias ORDER BY Categoria\"\n",
    b"    RsOpen rCat, \"SELECT DISTINCT Categoria FROM Categorias ORDER BY Categoria\"\n",
    "DISTINCT categorias"
)

# 2. Primeiro item das tags = "TODOS"
patch(
    b"    cboTagsF.AddItem \"\"\n"
    b"    If cboCategoriaF.Text = \"\" Then Exit Sub\n",
    b"    cboTagsF.AddItem \"TODOS\"\n"
    b"    If cboCategoriaF.Text = \"\" Then Exit Sub\n",
    "TODOS tags"
)

# 3. CarregarGrid: WHERE reformulado
patch(
    b"    ConfigurarGrid\n"
    b"    lstProdutos.Rows = 1\n"
    b"\n"
    b"    w = \"WHERE NCM IS NOT NULL AND NCM <> '' AND NCM <> '0'\"\n"
    b"    If cboCategoriaF.Text <> \"\" Then\n"
    b"        w = w & \" AND categoria = '\" & Replace(cboCategoriaF.Text, \"'\", \"''\") & \"'\"\n"
    b"    End If\n"
    b"    If cboTagsF.Text <> \"\" Then\n"
    b"        w = w & \" AND TAGS = '\" & Replace(cboTagsF.Text, \"'\", \"''\") & \"'\"\n"
    b"    End If\n"
    b"    If Trim(txtCodBarraF.Text) <> \"\" Then\n"
    b"        w = w & \" AND cod_barra LIKE '%\" & Replace(Trim(txtCodBarraF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n"
    b"    If Trim(txtDescF.Text) <> \"\" Then\n"
    b"        w = w & \" AND descricao LIKE '%\" & Replace(Trim(txtDescF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n"
    b"    If Trim(txtNCMF.Text) <> \"\" Then\n"
    b"        w = w & \" AND NCM LIKE '%\" & Replace(Trim(txtNCMF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n",
    b"    ConfigurarGrid\n"
    b"    lstProdutos.Rows = 1\n"
    b"\n"
    b"    ' Se algum campo de texto estiver preenchido, ignora filtros de combo\n"
    b"    Dim bTextSearch As Boolean\n"
    b"    bTextSearch = (Trim(txtCodBarraF.Text) <> \"\" Or Trim(txtDescF.Text) <> \"\" Or Trim(txtNCMF.Text) <> \"\")\n"
    b"\n"
    b"    w = \"WHERE 1=1\"\n"
    b"    If Not bTextSearch Then\n"
    b"        If cboCategoriaF.Text <> \"\" Then\n"
    b"            w = w & \" AND categoria = '\" & Replace(cboCategoriaF.Text, \"'\", \"''\") & \"'\"\n"
    b"        End If\n"
    b"        If cboTagsF.Text <> \"\" And cboTagsF.Text <> \"TODOS\" Then\n"
    b"            w = w & \" AND TAGS = '\" & Replace(cboTagsF.Text, \"'\", \"''\") & \"'\"\n"
    b"        End If\n"
    b"    End If\n"
    b"    If Trim(txtCodBarraF.Text) <> \"\" Then\n"
    b"        w = w & \" AND cod_barra LIKE '%\" & Replace(Trim(txtCodBarraF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n"
    b"    If Trim(txtDescF.Text) <> \"\" Then\n"
    b"        w = w & \" AND descricao LIKE '%\" & Replace(Trim(txtDescF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n"
    b"    If Trim(txtNCMF.Text) <> \"\" Then\n"
    b"        w = w & \" AND NCM LIKE '%\" & Replace(Trim(txtNCMF.Text), \"'\", \"''\") & \"%'\"\n"
    b"    End If\n",
    "WHERE logica"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
