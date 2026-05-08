"""
Patch Produtos_Cadastro.frm: campo TAGS
1. cboFabricante_GotFocus: adiciona sub cboTAGs_GotFocus logo depois
2. Carga ao editar: cboTAGs.Text = r("TAGS")
3. LimparObjetos: cboTAGs.Text = ""
4. INSERT coluna: TAGS antes de EANEmbalagem
5. INSERT valor: IIf TAGS antes de EANEmbalagem
6. UPDATE: TAGS apos tipo_calculo_is
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

# 1. Adicionar cboTAGs_GotFocus apos cboFabricante_GotFocus
patch(
    b"cboFabricante.Text = vTextoAntes\n"
    b"\n"
    b"moCombo.AttachTo cboFabricante\n"
    b"End Sub\n",
    b"cboFabricante.Text = vTextoAntes\n"
    b"\n"
    b"moCombo.AttachTo cboFabricante\n"
    b"End Sub\n"
    b"\n"
    b"\n"
    b"Private Sub cboTAGs_GotFocus()\n"
    b"Dim vTextoAntes As String\n"
    b"Dim rTags As ADODB.Recordset\n"
    b"\n"
    b"vTextoAntes = cboTAGs.Text\n"
    b"cboTAGs.Clear\n"
    b"RsOpen rTags, \"SELECT DISTINCT TAGS FROM produtos WHERE TAGS IS NOT NULL AND TAGS <> '' ORDER BY TAGS\"\n"
    b"Do While Not rTags.EOF\n"
    b"    cboTAGs.AddItem rTags(\"TAGS\")\n"
    b"    rTags.MoveNext\n"
    b"Loop\n"
    b"If rTags.State <> 0 Then rTags.Close\n"
    b"cboTAGs.Text = vTextoAntes\n"
    b"End Sub\n",
    "cboTAGs_GotFocus"
)

# 2. Carga ao editar
patch(
    b"    SelecionarNoCombo cboUnidMedida, ValidateNull(r(\"unid_medida\"))\n"
    b"\n"
    b"\n"
    b"End Sub\n",
    b"    SelecionarNoCombo cboUnidMedida, ValidateNull(r(\"unid_medida\"))\n"
    b"    cboTAGs.Text = ValidateNull(r(\"TAGS\"))\n"
    b"\n"
    b"\n"
    b"End Sub\n",
    "carga_editar"
)

# 3. LimparObjetos
patch(
    b"cboFabricante.Text = \"\"\n"
    b"'cboCategoria.Text = \"\"\n",
    b"cboFabricante.Text = \"\"\n"
    b"cboTAGs.Text = \"\"\n"
    b"'cboCategoria.Text = \"\"\n",
    "LimparObjetos"
)

# 4. INSERT coluna
patch(
    b"tipo_calculo_is, EANEmbalagem, Fracionamento) VALUES (\" & _\n",
    b"tipo_calculo_is, TAGS, EANEmbalagem, Fracionamento) VALUES (\" & _\n",
    "INSERT coluna"
)

# 5. INSERT valor
patch(
    b"IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\") & \", '\" & txtEANCaixa.Text",
    b"IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\") & \", \" & IIf(Trim(cboTAGs.Text) = \"\", \"NULL\", \"'\" & Replace(Trim(cboTAGs.Text), \"'\", \"''\") & \"'\") & \", '\" & txtEANCaixa.Text",
    "INSERT valor"
)

# 6. UPDATE
patch(
    b"    sSQL = sSQL & \"tipo_calculo_is = \" & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\")\n"
    b"\n"
    b"    ' --- Condi\xe7\xe3o Final ---\n",
    b"    sSQL = sSQL & \"tipo_calculo_is = \" & IIf(cboTipoIS.ListIndex >= 0, Val(Left(cboTipoIS.Text, 1)), \"NULL\") & \", \"\n"
    b"    sSQL = sSQL & \"TAGS = \" & IIf(Trim(cboTAGs.Text) = \"\", \"NULL\", \"'\" & Replace(Trim(cboTAGs.Text), \"'\", \"''\") & \"'\")\n"
    b"\n"
    b"    ' --- Condi\xe7\xe3o Final ---\n",
    "UPDATE TAGS"
)

print(f"\nTotal erros: {errors}")
data = data.replace(b"\n", b"\r\n")
with open(FILE, "wb") as f:
    f.write(data)
print("Arquivo gravado")
